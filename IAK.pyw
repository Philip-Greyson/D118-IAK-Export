"""Script to generate and send the IAK upload file via email.

https://github.com/Philip-Greyson/D118-IAK-Export

Does a real simple SQL query for staff, exports the info to a .csv file.
Then sends the .csv file via email to a specified email address.

Needs the google-api-python-client, google-auth-httplib2 and the google-auth-oauthlib:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
also needs oracledb: pip install oracledb --upgrade
"""

import os
from datetime import *
import oracledb
# import google API libraries
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
# libraries needed for emailing
import base64
import mimetypes
from email.message import EmailMessage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

# setup db connection
DB_UN = os.environ.get('POWERSCHOOL_READ_USER')  # username for read-only database user
DB_PW = os.environ.get('POWERSCHOOL_DB_PASSWORD')  # the password for the database account
DB_CS = os.environ.get('POWERSCHOOL_PROD_DB')  # the IP address, port, and database name to connect to
print(f'DBUG: Database Username: {DB_UN} |Password: {DB_PW} |Server: {DB_CS}')  # debug so we can see where oracle is trying to connect to/with

OUTPUT_FILE_NAME = 'iak_user_list.csv'
EMAIL_TARGET = os.environ.get('D118_IT_EMAIL')  # get the email to send to from the environment variable
SUB_BUILDING_CODE = 500  # the building code for substitutes in PowerSchool
BAD_NAMES = ['use', 'training1','trianing2','trianing3','trianing4','planning','admin','nurse','user','use ','test','testtt','do not','do','not','tbd','lunch','new','teacher','new teacher','teacher-1','sub','substitute','plugin','mba','tech','technology','administrator']  # List of names that some of the dummy/old accounts use so we can ignore them

# Google API Scopes that will be used. If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

if __name__ == '__main__':
    with open('iak_log.txt', 'w') as log:
        with open(OUTPUT_FILE_NAME, 'w') as output:
            startTime = datetime.now()
            startTime = startTime.strftime('%H:%M:%S')
            print(f'INFO: Execution started at {startTime}')
            print(f'INFO: Execution started at {startTime}', file=log)
            print('CARD ID,BADGE NAME,LAST NAME,FIRST NAME,MIDDLE NAME,FIELD1,FIELD2,FIELD3,FIELD4,FIELD5,FIELD6,FIELD7,FIELD8,SCHOOL GROUP NAME,RECORD ID',file=output)  # print out the header row

            # Do the authentication to Google API services
            # Get credentials from json file, ask for permissions on scope or use existing token.json approval, then build the "service" connection to Google API
            creds = None
                # The file token.json stores the user's access and refresh tokens, and is
                # created automatically when the authorization flow completes for the first
                # time.
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())

            service = build('gmail', 'v1', credentials=creds)

            # create the connecton to the PowerSchool database
            with oracledb.connect(user=DB_UN, password=DB_PW, dsn=DB_CS) as con:
                with con.cursor() as cur:  # start an entry cursor
                    cur.execute('SELECT dcid, teachernumber, last_name, first_name, homeschoolid FROM users WHERE email_addr IS NOT NULL ORDER BY last_name')
                    users = cur.fetchall()
                    for user in users:
                        try:
                            badgeType = 'Staff'  # set badge type to staff by default, will get overwritten to Sub below
                            dcid = user[0]
                            teacherNum = int(user[1])
                            lastName = str(user[2])
                            firstName = str(user[3])
                            homeschool = int(user[4])
                            if lastName.lower() not in BAD_NAMES and firstName.lower() not in BAD_NAMES:  # filter out a lot of the test/utility accounts
                                cur.execute('SELECT schoolid, staffstatus, status FROM schoolstaff WHERE users_dcid = :dcid', dcid=dcid)
                                schools = cur.fetchall()
                                for school in schools:
                                    schoolCode = int(school[0])
                                    staffType = int(school[1])
                                    schoolActive = True if school[2] == 1 else False
                                    if (schoolCode == SUB_BUILDING_CODE and schoolActive) or (staffType == 4):  # if they are active in the sub building or they have a stafftype code of substitute, override the badge type 
                                        badgeType = 'Substitute'
                                    if schoolCode == homeschool:  # check to see if the current school is their homeschool
                                        staffActive = schoolActive  # their active status at their homeschool is assumed to be their overall staff active status
                                if staffActive:  # only print them out if they are marked as active at their homeschool
                                    print(f'{teacherNum},{badgeType},{lastName},{firstName},,,,,,,,,,All,')
                                    print(f'{teacherNum},{badgeType},{lastName},{firstName},,,,,,,,,,All,', file=output)
                        except Exception as er:
                            print(f'ERROR while processing user with DCID {user[0]}: {er}')
                            print(f'ERROR while processing user with DCID {user[0]}: {er}', file=log)


        # # Create and send the email
        mime_message = EmailMessage() # create a email message object
        # # headers
        mime_message['To'] = EMAIL_TARGET # the email address it is sent to
        mime_message['Subject'] = 'Ident-A-Kid User List For ' + datetime.now().strftime('%Y-%m-%d') # subject line of the email, change to your liking
        # mime_message.set_content(f"Warning, there were {errorCount} errors in recent scripts: \n{errorString}") # the body of the email, aka the text

        # # attachment
        attachment_filename = OUTPUT_FILE_NAME # tell the email what file we are attaching
        # # guessing the MIME type
        type_subtype, _ = mimetypes.guess_type(attachment_filename)
        maintype, subtype = type_subtype.split('/')

        with open(attachment_filename, 'rb') as fp:
            attachment_data = fp.read() # read the file data in and store it in the attachment_data
        mime_message.add_attachment(attachment_data, maintype, subtype, filename=OUTPUT_FILE_NAME) # add the attacment data to the message object, give it a filename that was our output name

        # # encoded message
        encoded_message = base64.urlsafe_b64encode(mime_message.as_bytes()).decode()
        create_message = {
            'raw': encoded_message
        }
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
        print(f'INFO: Email sent, message ID: {send_message["id"]}') # print out resulting message Id
        print(f'INFO: Email sent, message ID: {send_message["id"]}', file=log)