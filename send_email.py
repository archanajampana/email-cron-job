# import smtplib
# import ssl
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# import os

# def send_email():
#     sender_email = os.getenv("SENDER_EMAIL")
#     sender_password = os.getenv("SENDER_PASSWORD")
#     receiver_email = os.getenv("RECEIVER_EMAIL")

#     subject = "Automated Email from GitHub Workflow"
#     body = "This is a test email sent every 5 minutes via GitHub Actions."

#     message = MIMEMultipart()
#     message["From"] = sender_email
#     message["To"] = receiver_email
#     message["Subject"] = subject
#     message.attach(MIMEText(body, "plain"))

#     context = ssl.create_default_context()
#     with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
#         server.login(sender_email, sender_password)
#         server.sendmail(sender_email, receiver_email, message.as_string())

# if __name__ == "__main__":
#     send_email()

import win32com.client
from datetime import datetime

def send_automation_status():
    try:
        # Create Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # Email details
        mail.To = "abbireddy.niharika@maersk.com"
        mail.CC = "kanta.rohitha@maersk.com"
        mail.Subject = f"Automation Run Status - {datetime.now():%Y-%m-%d %H:%M}"

        # HTML body for better formatting
        mail.HTMLBody = f"""
        <html>
        <body style="font-family: Calibri, Arial, sans-serif;">
            <p>Hello Team,</p>
            <p>The daily automation job has completed successfully at 
            <b>{datetime.now().strftime('%H:%M:%S')}</b>.</p>
            
            <p><b>Summary:</b></p>
            <ul>
                <li>Records Processed: 120</li>
                <li>Successful: 118</li>
                <li>Failed: 2</li>
            </ul>

            <p>Regards,<br>Automation Bot</p>
        </body>
        </html>
        """

        mail.Send()
        print("✅ Email sent successfully!")

    except Exception as e:
        print(f"❌ Error sending email: {e}")

if __name__ == "__main__":
    send_automation_status()

