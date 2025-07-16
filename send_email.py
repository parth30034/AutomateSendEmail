import os
import base64
import json
import time
import pandas as pd
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# If modifying the email, the scope needs to be adjusted
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Path to your credentials.json file
CREDENTIALS_FILE = 'credentials.json'

def authenticate_gmail():
    """Authenticate the user via OAuth 2.0 and return the service."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # Refresh if expired
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=8080)  # OAuth flow

        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def send_email(service, sender, to, first_name, company_name, job_id, attachment=None):
    """Send an email with an optional attachment using Gmail API."""
    try:
        # Create the MIME email
        message = MIMEMultipart()
        message['to'] = to
        message['subject'] = f""" Application for Data Engineer Role at {company_name} – Strong Fit with Data Engineering & Cloud Expertise """

        # Email body with dynamic first name
        body = f"""
        <html>
        <body>
            <p><b>Hi {first_name},</b></p>
            <p>I hope you are doing well.</p>
            <p>I recently applied for the Data Engineer position at {company_name} (Job Req ID <b>{job_id}</b>), and I wanted to take a moment to highlight my strong alignment with this role, particularly due to my background in data engineering and scalable cloud-based development.</p>
            
            <h3>Why I Am a Strong Fit for {company_name}:</h3>
            <ul>
                <li><strong>Professional Experience</strong> – I currently work as a <strong>Data Engineer</strong> at <strong>Mediaocean</strong>, where I focus on building and optimizing data pipelines, enhancing data processing workflows, and ensuring data integrity through automation and real-time analytics.</li>
                
                <li><strong>Technical Expertise in Data Engineering</strong> – Skilled in <strong>Python, PySpark, SQL</strong>, and Big Data tools like <strong>Apache Spark, Hadoop, and Kafka</strong>. I have extensive experience in designing and optimizing data pipelines and working with large datasets.</li>
                
                <li><strong>Cloud & Data Warehousing</strong> – Proficient in cloud platforms such as <strong>AWS (S3, EC2, Lambda, Glue, Redshift)</strong> and <strong>GCP (BigQuery)</strong>, along with data warehousing tools like <strong>Snowflake</strong> and <strong>MongoDB</strong>.</li>
                
                <li><strong>ETL Pipelines & Data Analytics</strong> – Expertise in building and managing <strong>ETL pipelines</strong> for data processing and analysis. Focused on ensuring data quality and implementing effective data governance.</li>
                
                <li><strong>DevOps & Automation</strong> – Hands-on experience with <strong>Docker, Jenkins, and CI/CD pipelines</strong>, driving automation and ensuring seamless cloud-based deployments.</li>
                
                <li><strong>Machine Learning & Data Testing</strong> – Knowledge of applying <strong>machine learning models</strong> to solve data challenges and leveraging automated testing frameworks for data integrity.</li>

                <li><strong>Education</strong>I completed my <strong>BTech in Computer Science and Engineering</strong> from Pimpri Chinchwad College of Engineering, Pune, with a CGPA of 9.57.</li>
            </ul>
            
            <p>Given my <b>technical expertise, problem-solving abilities, and hands-on experience in cloud-based data engineering,</b> I am confident I would bring valuable contributions to {company_name}. I would appreciate it if you could take a moment to review my application or let me know if you'd be open to discussing my qualifications further.</p>
            
            <p>Looking forward to hearing from you and exploring how I can contribute to your team.</p>
            
            <p>Best regards,<br>
            <b>Parth Shrivastava</b><br>
            Mob. No: 9860468498<br>
            Email Id: <a href="mailto:parth.r.shrivastava@gmail.com>parth.r.shrivastava@gmail.com/a><br>
            Linkedin: <a href="https://www.linkedin.com/in/parth-shrivastava-a68a78178/">https://www.linkedin.com/in/parth-shrivastava-a68a78178/</a></p>
        </body>
        </html>
        """

        # Attach the email body as HTML
        message.attach(MIMEText(body, 'html'))

        # Attach file if provided
        if attachment:
            with open(attachment, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={os.path.basename(attachment)}'
                )
                message.attach(part)

        # Encode the message to base64
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        message = {'raw': raw_message}
        message = service.users().messages().send(userId=sender, body=message).execute()
        print(f"Message Id: {message['id']}")
        return message
    except HttpError as error:
        print(f"An error occurred: {error}")


# def send_email(service, sender, to, first_name, company_name, job_id, attachment=None):
#     """Send an email with an optional attachment using Gmail API."""
#     try:
#         # Create the MIME email
#         message = MIMEMultipart()
#         message['to'] = to
#         message['subject'] = f""" Application for Data Engineer Roles at {company_name} – Strong Fit for Senior Data Engineer / Data Engineer """

#         # Email body with dynamic first name
#         body = f"""
#         <html>
#         <body>
#             <p><b>Hi {first_name},</b></p>
#             <p>I hope you are doing well.</p>
#             <p>I recently applied for the Data Engineer roles at  {company_name} (Job Req IDs <b>{job_id}</b>), and I wanted to take a moment to highlight my strong alignment with these positions.</p>
            
#             <h3>Why I Am a Strong Fit for {company_name}:</h3>
#             <ul>
#                 <li><strong>Professional Experience</strong> – I currently work as a <strong>Data Engineer</strong> at <strong>Mediaocean</strong>, where I focus on building and optimizing data pipelines, enhancing data processing workflows, and ensuring data integrity through automation and real-time analytics.</li>
                
#                 <li><strong>Technical Expertise in Data Engineering</strong> – Skilled in <strong>Python, PySpark, SQL</strong>, and Big Data tools like <strong>Apache Spark, Hadoop, and Kafka</strong>. I have extensive experience in designing and optimizing data pipelines and working with large datasets.</li>
                
#                 <li><strong>Cloud & Data Warehousing</strong> – Proficient in cloud platforms such as <strong>AWS (S3, EC2, Lambda, Glue, Redshift)</strong> and <strong>GCP (BigQuery)</strong>, along with data warehousing tools like <strong>Snowflake</strong> and <strong>MongoDB</strong>.</li>
                
#                 <li><strong>ETL Pipelines & Data Analytics</strong> – Expertise in building and managing <strong>ETL pipelines</strong> for data processing and analysis. Focused on ensuring data quality and implementing effective data governance.</li>
                
#                 <li><strong>DevOps & Automation</strong> – Hands-on experience with <strong>Docker, Jenkins, and CI/CD pipelines</strong>, driving automation and ensuring seamless cloud-based deployments.</li>
                
#                 <li><strong>Machine Learning & Data Testing</strong> – Knowledge of applying <strong>machine learning models</strong> to solve data challenges and leveraging automated testing frameworks for data integrity.</li>

#                 <li><strong>Education</strong> I completed my <strong>BTech in Computer Science and Engineering</strong> from Pimpri Chinchwad College of Engineering, Pune, with a CGPA of 9.57.</li>
#             </ul>
            
#             <p>Given my <b>technical expertise, problem-solving abilities, and hands-on experience in cloud-based data engineering,</b> I am confident I would bring valuable contributions to {company_name}. I would appreciate it if you could take a moment to review my application or let me know if you'd be open to discussing my qualifications further.</p>
            
#             <p>Looking forward to hearing from you and exploring how I can contribute to your team.</p>
            
#             <p>Best regards,<br>
#             <b>Parth Shrivastava</b><br>
#             Mob. No: 9860468498<br>
#             Email Id: <a href="mailto:parth.r.shrivastava@gmail.com">parth.r.shrivastava@gmail.com</a><br>
#             Linkedin: <a href="https://www.linkedin.com/in/parth-shrivastava51/">https://www.linkedin.com/in/parth-shrivastava51/</a></p>
#         </body>
#         </html>
#         """

#         # Attach the email body as HTML
#         message.attach(MIMEText(body, 'html'))

#         # Attach file if provided
#         if attachment:
#             with open(attachment, 'rb') as file:
#                 part = MIMEBase('application', 'octet-stream')
#                 part.set_payload(file.read())
#                 encoders.encode_base64(part)
#                 part.add_header(
#                     'Content-Disposition',
#                     f'attachment; filename={os.path.basename(attachment)}'
#                 )
#                 message.attach(part)

#         # Encode the message to base64
#         raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

#         message = {'raw': raw_message}
#         message = service.users().messages().send(userId=sender, body=message).execute()
#         print(f"Message Id: {message['id']}")
#         return message
#     except HttpError as error:
#         print(f"An error occurred: {error}")

def read_emails_from_excel(file_path):
    """Read email data from an Excel file."""
    df = pd.read_excel(file_path)
    email_data = []
    for _, row in df.iterrows():
        first_name = row['First Name']
        email = row['Email']
        company_name = row['Company Name']
        job_id = row['Job Id']

        email_data.append((first_name, email, company_name, job_id))
    return email_data

def main():
    sender = "parth.r.shrivastava@gmail.com"  # Replace with your email

    # Path to your Excel file with email data
    excel_file = 'emails.xlsx'
    # Path to the document you want to attach
    attachment_path = 'Resume_Parth_Shrivastava.pdf'  # Replace with your document's path

    # Authenticate and get the Gmail service
    service = authenticate_gmail()

    if service:
        email_data = read_emails_from_excel(excel_file)

        for first_name, email, company_name, job_id in email_data:
            send_email(service, sender, email, first_name, company_name, job_id, attachment_path)

if __name__ == '__main__':
    main()
