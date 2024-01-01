import win32com.client
import win32timezone
import pyodbc
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from win32com.client import constants
import re
import logging

logging.basicConfig(filename='service_log.txt', level=logging.DEBUG)
logging.info('Service started.')

db_connection_string = "DRIVER=ODBC Driver 17 for SQL Server;SERVER=10.11.12.13,1435;DATABASE=AMLMonitoring;UID=sa;PWD=benchmatrix786?"
# db_connection_string = "DRIVER=ODBC Driver 17 for SQL Server;SERVER=10.11.12.13,1435;DATABASE=AMLMonitoring;UID=sa;PWD=benchmatrix786?"
db_connection = pyodbc.connect(db_connection_string)
db_cursor = db_connection.cursor()

db_cursor.execute(
    """
    IF NOT EXISTS (
        SELECT * FROM sys.tables WHERE name='SupportEmailData'
    )
    CREATE TABLE SupportEmailData (
        ID INT IDENTITY(1,1) PRIMARY KEY,
        MessageID NVARCHAR(255),
        SenderName NVARCHAR(255),
        SenderEmail NVARCHAR(255),
        EmailDate DATETIME,
        Subject NVARCHAR(255),
        Body NVARCHAR(MAX),
        ClientName NVARCHAR(255)
    )
    """
)
db_connection.commit()

db_cursor.execute(
    """
    IF NOT EXISTS (
        SELECT * FROM sys.tables WHERE name='SupportEmailCheck'
    )
    CREATE TABLE SupportEmailCheck (
        ClientName NVARCHAR(255),
        Subject NVARCHAR(255),
        Configured INT
    )
    """
)
db_connection.commit()

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Use 6 for the Inbox folder
except Exception as e:
    logging.error(e)

# Get the 'Filtered' folder
filtered_folder = None
for folder in inbox.Folders:
    if folder.Name == 'AML Monitoring':
        filtered_folder = folder
        break
    
logging.info('connection with outlook made.')

if filtered_folder is not None:
    messages = filtered_folder.Items
    print(messages.count)
else:
    print("The 'AML Monitoring' folder does not exist.")
    logging.info('Folder in outlook doesnt exist.')
    

#This below code isnt needed now due to the above filtered folder on Outlook logic

# support_group_email = "support.aml@benchmatrix.com"
# services_group_email = "aml.services@benchmatrix.com"
# # Check if "aml.services@benchmatrix.com" in TO or CC of any email
# filtered_messages = messages.Restrict("[To] = '{0}' Or [CC] = '{0}' Or [To] = '{1}' Or [CC] = '{1}'".format(services_group_email, support_group_email))

# filtered_messages = []
# for message in messages:
#     for recip in message.recipients: 
#         print("=================")
#         # print(recip.Type)
#         # print(recip.AddressEntry.GetExchangeUser())
#         # if(recip == support_group_email or recip == services_group_email):
#         #     filtered_messages.append(message)
            
#     # recipients = message.To + ";" + message.CC
#     # if (support_group_email.lower() in recipients.lower() or services_group_email.lower() in recipients.lower()):
#     # filtered_messages.append(message)
          

for message in messages:
    try:
        #logic to store emails in db once and not duplicate emails
        db_cursor.execute(
            "SELECT COUNT(*) FROM SupportEmailData WHERE MessageID = ?", (message.EntryID,)
        )
        count = db_cursor.fetchone()[0]

        if count == 0:
            sender_name = message.SenderName
            sender_email = message.SenderEmailAddress
            email_date = message.ReceivedTime
            email_subject = message.Subject
            email_body = message.HTMLBody

            subject_parts = email_subject.split('|')
            if len(subject_parts) > 1:
                client_name = subject_parts[1].strip()
            else:
                client_name = None

            db_cursor.execute(
                """
                INSERT INTO SupportEmailData (MessageID, SenderName, SenderEmail, EmailDate, Subject, Body, ClientName)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (message.EntryID, sender_name, sender_email, email_date, email_subject, email_body, client_name),
            )
            
    except Exception as e:
        error=str(e)
        # print("Error processing email:", error)


#insert data logic into 2nd checkemail table here
rows_to_insert = [
    ("FMFB (Branch)", "Talend Job Results | FMFB (Branch)", None),
    ("FMFB (Branchless)", "Talend Job Results | FMFB (Branchless)", None),
    ("Ubank", "Talend Job results | Ubank", None),
    ("Ubank(Branch)", "Talend Job results | Ubank(Branch)", None),
    ("Ubank (BranchLess)", "Talend Job results | Ubank (BranchLess)", None),
    ("Tajeer", "Talend Job Results | Tajeer- Auto Upload Watchlist - Results", None),
    ("AlJabr", "Talend Job Results  | AlJabr - Auto Upload Watchlist - Results", None),
    ("Tajeer", "Talend Job Results | Tajeer-Service Logs - Results", None),
    ("Tasheel", "Talend Job Results | Tasheel- Auto Upload Watchlist - Results", None),
    ("Souq", "Talend Job Results | Funding Souq- Auto Upload Watchlist - Results", None),
    ("DCI", "Talend Job Results | DCI - Data Issue", None),
    ("NFT", "Talend Job Results | NFT- Auto Upload Watchlist - Results", None),
    ("FMFB", "Job Results And Monitoring | FMFB", None),
    ("Tabby", "Talend Job Monitoring| Tabby- Results", None),
    ("Tajeer", "Talend Job Monitoring| Tajeer - Results", None),
    ("Al Saqi", "Talend Job Monitoring| Al Saqi - Results", None),
    ("CFDB", "Talend Job Monitoring| CFDB- Results", None),
    ("Microcred", "Talend Job Monitoring| Microcred - Results", None),
    
]

tquery = "Select COUNT(*) from SupportEmailCheck"
db_cursor.execute(tquery)
tcount = db_cursor.fetchone()[0]

if(tcount==0):
    db_cursor.executemany(
    """
    INSERT INTO SupportEmailCheck (ClientName, Subject, Configured)
    VALUES (?, ?, ?)
    """,
    rows_to_insert)
    db_connection.commit()



today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

db_cursor.execute("SELECT * FROM SupportEmailData WHERE CONVERT(DATE, EmailDate) = ?", yesterday)
day_1_emails = db_cursor.fetchall()

db_cursor.execute("SELECT DISTINCT Subject FROM SupportEmailCheck")
distinct_subjects = [row.Subject for row in db_cursor.fetchall()]

for subject in distinct_subjects:
    # Check if any email in day-1 emails has a matching subject
    is_configured = any(email.Subject == subject for email in day_1_emails)

    if(is_configured):
        db_cursor.execute("UPDATE SupportEmailCheck SET Configured = ? WHERE Subject = ?", (is_configured, subject))
        db_connection.commit()
    else:
        db_cursor.execute("UPDATE SupportEmailCheck SET Configured = 0 WHERE Subject = ?", (subject))
        db_connection.commit()
         
        
#Print no data generated text for clients with configured flag = 0       
db_cursor.execute("SELECT ClientName FROM SupportEmailCheck WHERE Configured = 0")
not_configured_clients = db_cursor.fetchall()
not_generated_messages = "\n".join(f"No data generated for '{client.ClientName}' <br>" for client in not_configured_clients)
    
    
db_cursor.execute("SELECT * FROM SupportEmailData WHERE CONVERT(DATE, EmailDate) = ? and Subject LIKE '%task scheduler updates%'", yesterday)
email_records = db_cursor.fetchall()
# db_connection.close() 
print("fetched")

message = MIMEMultipart()
message['From'] = 'm.rashid@benchmatrix.com'
message['To'] = 'aml.services@benchmatrix.com'
message['Subject'] = f"AML Monitoring Task Scheduler Updates"


keywords = [           #to be removed
    "This is an auto-generated message",
    "It is to be notified that",
    "This communication",
    "You received this message because",
    "Disclaimer: This email and any files transmitted",
    "To unsubscribe from this group",
    "Tasheel- Auto Upload Watchlist - Results",
    "Tajeer-Service Logs - Results",
    "Disclaimer: This email is confidential"
]

email_body = ""
for record in email_records:
    if record.Subject == "Task scheduler updates ( Error )":
        message['X-Priority'] = '1'

    if record.Body.count("`") == 2:
        start_index = record.Body.find("`")
        end_index = record.Body.find("`", start_index + 1)
        required_body = record.Body[start_index + 1: end_index]
        email_body += f"{required_body}\n\n <br>"
    else:
        # email_body += f"{record.Body} <br><br>/n/n"
        # email_body_parts = email_body.split('--')
        # email_body = email_body_parts[0]
        # email_body += "\n\n"
        body = record.Body
        for keyword in keywords:
            index = body.find(keyword)
            if index != -1:
                # body = body[:index].strip()
                body = body[:index]
                break
        email_body += f"{body} <br><br>"
# email_body += not_generated_messages   #not tested
        
        
html_message = MIMEText(email_body, 'html')
message.attach(html_message)

smtp_server = 'smtp.gmail.com'
smtp_port = 587  
smtp_username = 'm.rashid@benchmatrix.com'
smtp_password = 'benchmarkmr11?'

with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.send_message(message)

print("sent")



# Get emails with subject '%Talend Job Results%'
# db_cursor.execute("SELECT * FROM SupportEmailData WHERE CONVERT(DATE, EmailDate) = ? and Subject LIKE '%Talend Job Results%'", today)
db_cursor.execute("SELECT * FROM SupportEmailData WHERE CONVERT(DATE, EmailDate) = ? and Subject LIKE '%talend job results%'", yesterday)
email_records = db_cursor.fetchall()
if email_records is not None:
    print("fetched1")

message = MIMEMultipart()
message['From'] = 'm.rashid@benchmatrix.com'
# message['To'] = 'aml.jobs@benchmatrix.ca'     #change later
message['To'] = 'aml.services@benchmatrix.com'
message['Subject'] = f"AML Monitoring Talend Job Results"

sentences = [
    "BranchLess Data import  Completed",
    "Data Import job executed successfully.",
    "Branch Data Import Completed",
    "Branch Data import Completed",
    "Branchless Data Import Completed",
    "Tasheel- Auto Upload Watchlist - Results",
    "Tajeer-Service Logs - Results"
]

email_body = ""
for record in email_records:
    body = record.Body
    for sent in sentences:
        if sent in body:
            body = body.replace(sent, "")
    for keyword in keywords:
        index = body.find(keyword)
        if index != -1:
            body = body[:index].strip()
            break

    subject_parts = record.Subject.split('|')
    if len(subject_parts) > 1:
        heading = subject_parts[1].strip()
        email_body += f"<b>{heading}</b>"
    email_body += f"{body}<br>"

html_message = MIMEText(email_body, 'html')
message.attach(html_message)

smtp_server = 'smtp.gmail.com'
smtp_port = 587  
smtp_username = 'm.rashid@benchmatrix.com'
smtp_password = 'benchmarkmr11?'

with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.send_message(message)

print("sent")



db_cursor.execute("SELECT * FROM SupportEmailData WHERE CONVERT(DATE, EmailDate) = ? and Subject LIKE '%talend job monitoring%' ", yesterday)
email_records = db_cursor.fetchall()
if email_records is not None:
    print("fetched2")

message = MIMEMultipart()
message['From'] = 'm.rashid@benchmatrix.com'
message['To'] = 'aml.services@benchmatrix.com'
message['Subject'] = f"AML Monitoring Job Results and Monitoring"

sentences = [
    "BranchLess Data import  Completed",
    "Data Import job executed successfully.",
    "Branch Data Import Completed",
    "Branch Data import Completed",
    "Branchless Data Import Completed",
    "Tasheel- Auto Upload Watchlist - Results",
    "Tajeer-Service Logs - Results"
]

email_body = ""
for record in email_records:
    body = record.Body
    for sent in sentences:
        if sent in body:
            body = body.replace(sent, "")
    for keyword in keywords:
        index = body.find(keyword)
        if index != -1:
            body = body[:index].strip()
            break

    subject_parts = record.Subject.split('|')
    if len(subject_parts) > 1:
        heading = subject_parts[1].strip()
        email_body += f"<b>{heading}</b> <br>"
    email_body += f"{body}<br>"

email_body += not_generated_messages   
html_message = MIMEText(email_body, 'html')
message.attach(html_message)

smtp_server = 'smtp.gmail.com'
smtp_port = 587  
smtp_username = 'm.rashid@benchmatrix.com'
smtp_password = 'benchmarkmr11?'

with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.send_message(message)
    
print("sent")

db_connection.close()