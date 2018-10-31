from apiclient import errors, discovery
import base64
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import html2text
import httplib2
import json
import oauth2client
from oauth2client import client, tools
import os
import pyodbc
import time

"""
user_file = os.path.join(path, 'users.txt')
error_file = os.path.join(path, 'error_log.txt')
log_file = os.path.join(path, 'logs.txt')

direct_ip = settings["Direct IP"]
"""
path = os.path.dirname(os.path.realpath(__file__))
settings_file = os.path.join(path, 'settings.txt')
with open(settings_file, 'r') as f:
	settings = json.load(f)

sender = "No-reply@uview.academy"

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'credentials.json'
APPLICATION_NAME = 'Gmail API Python Send Email'

def get_credentials():
	home_dir = os.path.expanduser('~')
	credential_dir = os.path.join(home_dir, '.credentials')
	if not os.path.exists(credential_dir):
		os.makedirs(credential_dir)
	credential_path = os.path.join(credential_dir, 'gmail-python-email-send.json')
	store = oauth2client.file.Storage(credential_path)
	credentials = store.get()
	if not credentials or credentials.invalid:
		flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
		flow.user_agent = APPLICATION_NAME
		credentials = tools.run_flow(flow, store)
		print('Storing credentials to ' + credential_path)
	return credentials

def SendMessage(sender, to, subject, msgHtml, msgPlain):
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('gmail', 'v1', http=http)
	message1 = CreateMessage(sender, to, subject, msgHtml, msgPlain)
	SendMessageInternal(service, "me", message1)

def SendMessageInternal(service, user_id, message):
	try:
		message = (service.users().messages().send(userId=user_id, body=message).execute())
		print('Message Id: %s' % message['id'])
		return message
	except errors.HttpError as error:
		print('An error occurred: %s' % error)

def CreateMessage(sender, to, subject, msgHtml, msgPlain):
	msg = MIMEMultipart('alternative')
	msg['Subject'] = subject
	msg['From'] = sender
	msg['To'] = to
	msg.attach(MIMEText(msgPlain, 'plain'))
	msg.attach(MIMEText(msgHtml, 'html'))
	raw = base64.urlsafe_b64encode(msg.as_bytes())
	raw = raw.decode()
	body = {'raw': raw}
	return body


messages_to_send = {}

print("Connecting to database.")
cnxn = pyodbc.connect('Driver={SQL Server};' + 'Server={};Database={};Port={};UID={};PWD={}'.format(settings["Server"], settings["Database"], settings["Port"], settings["User"], settings["Password"]))
cnxn.setencoding('utf-8')
cursor = cnxn.cursor()

cursor.execute("SELECT CONCAT(st.LastName, ', ', st.FirstName) AS StudentName, st.GradeLevel, Fail.FailCount, sec.Name, sec.TeacherIndex, Users.Email FROM Students st INNER JOIN Enrollments e ON st.StudentIndex = e.StudentIndex INNER JOIN Sections sec ON sec.SectionIndex = e.SectionIndex INNER JOIN LMSTerms term ON term.LMSTermIndex = sec.LMSTermIndex INNER JOIN Teachers ON Teachers.TeacherIndex = sec.TeacherIndex INNER JOIN Users ON Users.UserIndex = Teachers.UserIndex INNER JOIN (SELECT st.StudentIndex, count(sec.Name) AS FailCount FROM Students st INNER JOIN Enrollments e ON st.StudentIndex = e.StudentIndex INNER JOIN Sections sec ON sec.SectionIndex = e.SectionIndex INNER JOIN LMSTerms term ON term.LMSTermIndex = sec.LMSTermIndex WHERE st.Status = 'active' AND e.Status = 'active' AND ((sec.CourseTypeCodeListIndex NOT IN (8, 9, 10) AND e.CurrentGrade < 67) OR ((sec.CourseTypeCodeListIndex IN (8, 9, 10) OR sec.Name LIKE '%%develop%%') AND e.CurrentGrade < 60)) AND st.GradeLevel IN ('9', '9t', '10', '11', '12') AND sec.Name NOT LIKE '%%advisory%%' AND term.SemesterIndex IN (1, 2) GROUP BY st.StudentIndex) Fail ON Fail.StudentIndex = st.StudentIndex WHERE st.Status = 'active' AND e.Status = 'active' AND sec.Name LIKE '%%ec advisory%%' AND st.GradeLevel IN ('9', '9t', '10', '11', '12') ORDER BY sec.Name, StudentName")
rows = cursor.fetchall()

for row in rows:
	try:
		messages_to_send[row[5]] += "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(row[0], row[1], row[2])
	except:
		messages_to_send[row[5]] = "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(row[0], row[1], row[2])

today = datetime.datetime.today().strftime('%m-%d')
subject = "Advisory Students Failing Courses ({})".format(today)
for x in messages_to_send:
	to = x
	message_html = "The following is a current summary of your advisory students:<br><br><table style='width:90%%;border-spacing:10px''><col align='left'><col align='center'><col align='center'><tr><th>Student Name</th><th>Grade Level</th><th>Number of Classes Currently Failing</th></tr>{}</table>".format(messages_to_send[x])
	message_plain = html2text.html2text(message_html)
	SendMessage(sender, to, subject, message_html, message_plain)
	time.sleep(1)

cursor.close()
del cursor
cnxn.close()

print("Connection closed.")

print(str(datetime.datetime.now()) + " Done.")