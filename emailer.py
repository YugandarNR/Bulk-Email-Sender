import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

# Configure smtp gmail server
smtp_server='smtp.gmail.com'
smtp_port = 587

# Create a custom email template
# Path to all the files
file_path=''

# Excel file name with reciever filenames and first names
excel_filename="emails.xlsx"

# sender email id
sender_email=''

# sender app password
sender_passcode=''

# Attachment filename
pdf_filename = 'pdf_test.pdf'

# Email Subject
email_subject ='This is a Test Email 7'

#Email template. Write \n to move to next line
email_template_plain = 'Hello {first_name},\nThis is a test email. Do not respond!\nThank you\nSun N\nemail: abcd@gmail.com\nphone: +1 (987) 654-3210'
email_template_html = """
<html>
  <body">
    <p>Hello {first_name},</p>
    <p style="margin: 0;">This is a test email. Do not respond!</p>
    <p style="margin-top: 0;">Please find link to my <a href="https://drive.google.com/file/d/1XsJl9vLONbf1Lr-W6-WckeEfrdcz6hrI/view?usp=drive_link">Resume</a></p>
    <p style="margin: 0;">Thank you</p>
    <p style="margin: 0;"></p>
    <p style="margin: 0;">Sun N</p>
    <p style="margin: 0;">email: <a href="mailto:abcd@gmail.com">abcd@gmail.com</a></p>
    <p style="margin: 0;">phone: +1 (987) 654-3210</p>
    
  </body>
</html>
"""

# Select which template to use
email_template=email_template_html
# use 'html' if using html template or else use 'plain'
email_template_type='html'


# Load the Excel file
wb = openpyxl.load_workbook(os.path.join(file_path,excel_filename))
sheet = wb['Sheet1']

# Login to the gmail server
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_passcode)
    print("Login successful!")
except smtplib.SMTPAuthenticationError:
    print("Login failed. Check your email and password.")
except smtplib.SMTPException as e:
    print("Login failed. Error:", e)

# Function to attach pdf
def attach_pdf(pdf_file_path,pdf_filename):
    with open(os.path.join(pdf_file_path,pdf_filename), 'rb') as f:
        pdf_attachment = MIMEBase('application', 'pdf')
        pdf_attachment.set_payload(f.read())
        encoders.encode_base64(pdf_attachment)
        pdf_attachment.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
        return pdf_attachment

pdf_attachment=attach_pdf(file_path,pdf_filename)

# Loop through the rows
for row in sheet.iter_rows(values_only=True):
    receiver_email = row[0]
    receiver_first_name = row[1]
    # Skip execution if receiver email or reeiver first name is not available
    if receiver_email is None or receiver_first_name is None:
        continue
    # Replace {firstname} with actual values
    email_content = email_template.replace('{first_name}', receiver_first_name)
    # Create Multipart Email
    msg = MIMEMultipart()
    msg.attach(MIMEText(email_content, email_template_type))
    msg.attach(pdf_attachment)
    msg['Subject'] = email_subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    try:
        server.sendmail(sender_email, receiver_email, msg.as_string())
        print(f"Email sent to {receiver_first_name} at {receiver_email}")
    except smtplib.SMTPRecipientsRefused as e:
        print(f"Error sending email to {receiver_email}: {e}")
    except smtplib.SMTPException as e:
        print(f"Error sending email: {e}")
server.quit()
