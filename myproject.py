import gspread
import cv2
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Load the Google Sheets API credentials
gc = gspread.service_account(filename=r"C:\Users\Dell\Downloads\hitesh-411510-62ba6eca1bbd.json")


# Open the Google Sheets document by title
doc = gc.open("Data")

# Assuming you have a sheet named 'Sheet1'; adjust if needed
worksheet = doc.get_worksheet(0)

# Read data from the worksheet
df = pd.DataFrame(worksheet.get_all_records())

# SMTP email configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USERNAME = 'hiteshgehlot803@gmail.com'
SMTP_PASSWORD = 'fvle nfyr swow xrch'

# List to store indices of rows to be deleted
rows_to_delete = []

for index, (name1, email) in enumerate(zip(df['Name'], df['Email'])):
    try:
        # Load the certificate image
        image = cv2.imread(r"C:\Users\Dell\Desktop\python 2023\certificate.png")

        # Add text to the certificate
        cv2.putText(image, name1, (650, 780), cv2.FONT_HERSHEY_COMPLEX, 3, (0, 0, 0), 1, cv2.LINE_AA)
        #cv2.putText(image, email, (730, 1170), cv2.FONT_HERSHEY_COMPLEX, 2, (0, 0, 0), 1, cv2.LINE_AA)

        # Save the certificate with a unique name
        certificate_path = r'C:\Users\Dell\Desktop\python 2023\{}_certificate.png'.format(name1)
 
        cv2.imwrite(certificate_path, image)

        # Send the certificate via email
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)

        msg = MIMEMultipart()
        msg['From'] = SMTP_USERNAME
        msg['To'] = email
        msg['Subject'] = 'Certificate of Completion'

        body = f'Dear {name1},\n\nCongratulations on completing your course! Please find your certificate attached.'
        msg.attach(MIMEText(body, 'plain'))

        with open(certificate_path, 'rb') as img_file:
            image = MIMEImage(img_file.read(), name=f'{name1}_certificate.png')
            msg.attach(image)

        server.sendmail(SMTP_USERNAME, email, msg.as_string())
        server.quit()

        print(f'Processing Certificate {index + 1}/{len(df)} - Sent to {email}')

        # Add the index to the list for later deletion
        rows_to_delete.append(index + 2)  # Assuming index starts from 0 and header is in the first row

    except Exception as e:
        print(f'Error sending certificate to {email}: {str(e)}')

    # Delete the temporary certificate file
    os.remove(certificate_path)

# Delete the rows from the worksheet after the email sending loop
for row_index in reversed(rows_to_delete):
    worksheet.delete_rows(row_index)