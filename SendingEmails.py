import smtplib
import pandas as pd
from datetime import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# -----------------------------
# Gmail Credentials
# -----------------------------
my_email = "sdeeksha906@gmail.com"
my_password = "gecn eudn kmyw dmve"

# -----------------------------
# Read CSV
# -----------------------------
df = pd.read_csv("DataSet.csv")

# -----------------------------
# Read Email Body Template
# -----------------------------
with open("Mail Body.txt", "r", encoding="utf-8") as file:
    template = file.read()

# -----------------------------
# Add Attachment
# -----------------------------
attachment =r"C:\Users\sdeek\Desktop\PYTHON\Mail on Gmail\CV_DeekshaSharma.docx"

# -----------------------------
# Current Date & Time
# -----------------------------
now = dt.now()
today_date = now.strftime("%d-%m-%Y")
current_time = now.strftime("%H:%M")

print("Current date:", today_date)
print("Current time:", current_time)

# ----------------------------------------
# Email Sending Function (with attachment)
# ----------------------------------------
def send_email(to_address, subject, body, attachment_path = attachment):
    msg = MIMEMultipart()
    msg["From"] = my_email
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # -----------------------------
    # Attachment Section
    # -----------------------------
    if attachment_path:
        with open(attachment_path, "rb") as attached:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attached.read())

        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f"attachment; filename={attachment_path.split('/')[-1]}")
        msg.attach(part)

    # -----------------------------
    # Send Email
    # -----------------------------
    with smtplib.SMTP("smtp.gmail.com") as connection:
        connection.starttls()
        connection.login(my_email, my_password)
        # Build UTF-8 encoded email
        message = (
            f"From: {my_email}\n"
            f"To: {to_address}\n"
            f"Subject: {subject}\n"
            f"Content-Type: text/plain; charset=utf-8\n\n"
            f"{body}"
        )

        connection.sendmail(my_email, to_address, message.encode("utf-8"))


# -----------------------------
# Process each row in CSV
# -----------------------------
for index, row in df.iterrows():
    name = row["Name"]
    email = row["Email"]
    date_val = row["Date"]
    time_val = row["Time"]

    # Check if date & time match the current system date/time
    if today_date == date_val and current_time == time_val:

        # Replace [Name] placeholder
        mail_body = template.replace("[Name]", name)

        print(f"Sending email to {name} at {email}...")

        send_email(
            to_address=email,
            subject="Python Automation Test",
            body=mail_body,
            attachment_path = attached
        )

        print("Email sent successfully!\n")
    else:
        print(f"Skipping {name} — time/date not matched.")