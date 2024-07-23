import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText

st.title("Automated Email Sender")

# Configuration
email_server = 'smtp.gmail.com'
email_port = 587

email_username = st.text_input("Enter email username")
email_password = st.text_input("Enter email password", type="password")

excel_file = st.file_uploader("Select Excel file", type=["xlsx"])

if excel_file and email_username and email_password:
    df = pd.read_excel(excel_file)

    passing_score = st.number_input("Enter passing score", min_value=0, value=60)

    for index, row in df.iterrows():
        student_name = row["Student Name"]
        student_score = row["Score"]
        student_email = row["Email"]
        college_dean = row["Dean"]
        college_name = row["College Name"]
        college_city = row["City"]

        if student_score >= passing_score:
            email_body = f'Dear Parent,  \n This is to inform you that your ward, {student_name} is eligible to appear for the Final Examination of CHSE This year. We wish your ward all of the best for their exam, and in future endeavors. We advise your ward to study sincerely for the examination so that they do well and bring laurels to our institution. Our professors and Faculty are now interacting with students through extra classes so that their needs can be better identified and addressed, and we request that your ward attend the same. \nThe student has our full support for these exams, seeing that this is an important part of their careers. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}'

            subject = 'You Have Passed!'
        else:
            email_body = f"'Dear Parent,  \nwe regret to inform you that your ward, {student_name}, has been declared ineligible for the upcoming CHSE Examinations. However, we expect that he will appear for the exam next year with renewed vigour, and we wish to inform you that the administration and faculty of the institution will leave no stone unturned in ensuring his success next year. \nWe would recommend for you to meet {student_name}'s professors, so that you may better understand his needs, requirements and shortcomings. \nWe hope you understand the mental pressure {student_name} is going through, and we urge you not to reprimand him at present. His success is our duty, and we shall accomplish it by all means we can. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}"

            subject = 'You Have Not Passed'

        # Send email
        msg = MIMEText(email_body)
        msg['Subject'] = subject
        msg['From'] = email_username
        msg['To'] = student_email

        server = smtplib.SMTP(email_server, email_port)
        server.starttls()
        server.login(email_username, email_password)
        server.sendmail(email_username, student_email, msg.as_string())
        server.quit()

        st.write(f"Email sent to {student_name} at {student_email}")
