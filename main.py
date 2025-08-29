import pandas as pd
import smtplib as sm
import socket
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Read the Excel file
data = pd.read_excel("Book1.xlsx")
email_col = data.get("email")
list_of_emails= list(email_col)
print(list_of_emails)

try:
    server = sm.SMTP('smtp.gmail.com', 587)
    server.set_debuglevel(1)
    server.starttls()
    server.login("testerrr649@gmail.com","twxbczkipqljyrad")
    from_add = "testerrr649@gmail.com"
    to = list_of_emails
    subject = "This is a test email"
    body = "mewww mew mew mew mew mew mew mew mew ."

#create message
    msg = MIMEMultipart()
    msg['From'] = from_add
    msg['To'] = ", ".join(to)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach photo
    filename = "mew.jpeg"  # Change to your file name
    with open(filename, "rb") as attachment:    
        mime_base = MIMEBase("application", "octet-stream")
        mime_base.set_payload(attachment.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(mime_base)

    # Send email
    server.sendmail(from_add, to, msg.as_string())  
    print("Email sent successfully!")

except Exception as e:
    print(f"Failed to send email. Error: {e}")
finally:
    server.quit()   