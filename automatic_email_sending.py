import smtplib
import configparser
from datetime import datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


def attach_file(*, message, path):
    if path:
        # Attach the file
        try:
            with open('Archive.csv', 'rb') as attachment:
                # Create a MIMEBase object
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                # Encode the payload using Base64
                encoders.encode_base64(part)
                # Add header to the part
                part.add_header('Content-Disposition', f'attachment; filename={'Archive.csv'}')
                # Attach the part to the message
                message.attach(part)
        except Exception as e:
            print(f"Error attaching file: {e}")


def send_daily_email(*, config, msg):
    try:
        # Connect to the SMTP server
        server = smtplib.SMTP(config['Settings']['smtp_server'], config['Settings'].getint('port'))
        server.starttls()  # Secure the connection
        server.login(config['Credentials']['username'], config['Credentials']['password'])

        # Send the email to each recipient
        for recipient in config['Other']['to'].split():
            msg['To'] = recipient
            server.sendmail(config['Credentials']['username'], recipient, msg.as_string())
            print(f"Email sent successfully to {recipient}")

        # Quit the SMTP server
        server.quit()

    except Exception as e:
        print(f"Error sending email: {e}")


def main():
    # Current date
    today = dt.now().strftime('%d.%m.%y')

    # Taking data from config.ini file
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = config['Credentials']['username']
    msg['To'] = config['Other']['to']
    msg['Subject'] = f'BATYSPETROLEUM: Архив событий за {today}'

    # Email body
    body = config['Other']['email_body']
    msg.attach(MIMEText(body, 'plain'))

    # Attaching the file and sending the email to the recipients
    attach_file(message=msg, path=config['Other']['attachment_path'])
    send_daily_email(config=config, msg=msg)


if __name__ == '__main__':
    main()
