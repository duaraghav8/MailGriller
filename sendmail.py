#!/usr/bin/env python3
import smtplib;
import os;
from openpyxl import load_workbook;

#For problems, feedback and/or bug reporting, please contact duaraghav8@gmail.com
#Possible improvements: instead of storing password in plaintext in the script (dangerous!), we could store it in form of cyphertext. We could alternatively ask for password each time the script is run.

def send_mail (to, subject, text, gmail_sender_id, server):
	BODY = '\r\n'.join(['To: %s' % to, 'From: %s' % gmail_sender_id,'Subject: %s' % subject, '', text]);

	try:
		server.sendmail(gmail_sender_id, [to], BODY);
		print ('email sent to ', to);
	except:
		print ('error sending mail to ', gmail_sender_id);

def get_recipients (excel_file):
	data_file = load_workbook (excel_file);
	sheet = data_file.active;
	recipient_data = {};
	name = '';
	email = 'xxxx@xxxx.xxx';
	columns = ['A', 'B'];
	counter = 1;	#This assumes that the data begins from Row 1. If row 1 is your titles and the actual data begins from Row 2, then set counter to 2 instead of 1

	while (True):
		email_index = columns [0] + str (counter);
		name_index = columns [1] + str (counter);

		email = sheet [email_index].value;
		if (not email):
			break;
		name = sheet [name_index].value;

		recipient_data [email] = name;
		counter += 1;
	return (recipient_data);

def initialize_server (gmail_sender_id, gmail_sender_passwd):
	server = smtplib.SMTP ('smtp.gmail.com', 587);
	server.ehlo ();
	server.starttls ();
	server.login (gmail_sender_id, gmail_sender_passwd);
	return (server);

if (__name__ == '__main__'):
	mail_recipients = get_recipients ('recipients.xlsx');
	print (mail_recipients);
#	Gmail Account Sign In: Enter the User ID and password below, so it would seem that you're sending an email from this account
	gmail_sender_id = '';
	gmail_sender_passwd = '';
	hotspot = '****';

#	TYPE SUBJECT and give the name of the text file which contains the mail's body
	subject = input ('Subject: ');
	mailfile = input ('Enter name of file from which you wish to send the message body: ');
	text = open (mailfile).read ();

	server = initialize_server (gmail_sender_id, gmail_sender_passwd);

	confirm = input ('Everything set. Go? (y / n): ');
	if (not confirm == 'y'):
		os._exit (1);

	for recipient in list (mail_recipients.keys ()):
		name = mail_recipients [recipient];
		text = text.replace (hotspot, name);
		send_mail (recipient, subject, text, gmail_sender_id, server);

	server.quit ();
	print ('\nThank for using sendmail service. PS You know the guy who created this for you? Give that man a cookie!');
