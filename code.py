import streamlit as st
import pandas as pd
import openpyxl
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

st.title("Automated Email Sender for Students")

# Configuration
passing_score = 60
sendgrid_api_key = 'YOUR_SENDGRID_API_KEY'

# Load Excel file
wb = openpyxl.load_workbook('C:\\Users\\91891\\OneDrive\\Desktop\\Book1.xlsx')
sheet = wb.active
rows = sheet.iter_rows(values_only=True)
next(rows)  # Skip the header row

# Create a button to send emails
if st.button("Send Emails"):
    for row in rows:
        student_name = row[0]
        student_score = int(row[1])
        student_email = row[2]
        college_dean = row[3]
        college_name = row[4]
        college_city = row[5]

        if student_score >= passing_score:
            email_body = f'Dear Parent,  \n This is to inform you that your ward, {student_name} is eligible to appear for the Final Examination of CHSE This year. We wish your ward all of the best for their exam, and in future endeavors. We advise your ward to study sincerely for the examination so that they do well and bring laurels to our institution. Our professors and Faculty are now interacting with students through extra classes so that their needs can be better identified and addressed, and we request that your ward attend the same. \nThe student has our full support for these exams, seeing that this is an important part of their careers. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}'

            subject = 'You Have Passed!'
        else:
            email_body = f"'Dear Parent,  \nwe regret to inform you that your ward, {student_name}, has been declared ineligible for the upcoming CHSE Examinations. However, we expect that he will appear for the exam next year with renewed vigour, and we wish to inform you that the administration and faculty of the institution will leave no stone unturned in ensuring his success next year. \nWe would recommend for you to meet {student_name}'s professors, so that you may better understand his needs, requirements and shortcomings. \nWe hope you understand the mental pressure {student_name} is going through, and we urge you not to reprimand him at present. His success is our duty, and we shall accomplish it by all means we can. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}"

            subject = 'You Have Not Passed'

        # Send email using SendGrid
        message = Mail(
            from_email='your_email@example.com',
            to_emails=student_email,
            subject=subject,
            plain_text_content=email_body
        )
        try:
            sg = SendGridAPIClient(sendgrid_api_key)
            response = sg.send(message)
            st.write(f'Email sent to {student_name} at {student_email}')
        except Exception as e:
            st.write(f'Error sending email to {student_name} at {student_email}: {e}')
