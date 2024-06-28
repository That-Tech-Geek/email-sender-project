import smtplib
from email.mime.text import MIMEText
import openpyxl
passing_score = 60
email_server = 'smtp.gmail.com'
email_port = 587
email_username = 'email of user'
email_password = 'insert app password to allow access'
wb = openpyxl.load_workbook('C:\\Users\\91891\\OneDrive\\Desktop\\Book1.xlsx')
sheet = wb.active
rows = sheet.iter_rows(values_only=True)
next(rows)  # Skip the header row
for row in rows:
    student_name = row[0]
    student_score = int(row[1])
    student_email = row[2]
    college_dean = row[3]
    college_name = row[4]
    college_city = row[5]
    if student_score >= passing_score:
        # Define the email body with the college dean's name
        email_body = f'Dear Parent,  \n This is to inform you that your ward, {student_name} is eligible to appear for the Final Examination of CHSE This year. We wish your ward all of the best for their exam, and in future endeavors. We advise your ward to study sincerely for the examination so that they do well and bring laurels to our institution. Our professors and Faculty are now interacting with students through extra classes so that their needs can be better identified and addressed, and we request that your ward attend the same. \nThe student has our full support for these exams, seeing that this is an important part of their careers. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}'

        # Send an email to the student
        msg = MIMEText(email_body)
        msg['Subject'] = 'You Have Passed!'
        msg['From'] = email_username
        msg['To'] = student_email

        server = smtplib.SMTP(email_server, email_port)
        server.starttls()
        server.login(email_username, email_password)
        server.sendmail(email_username, student_email, msg.as_string())
        server.quit()

        print(f'Email sent to {student_name} at {student_email}')
    else:
        email_body = f"'Dear Parent,  \nwe regret to inform you that your ward, {student_name}, has been declared ineligible for the upcoming CHSE Examinations. However, we expect that he will appear for the exam next year with renewed vigour, and we wish to inform you that the administration and faculty of the institution will leave no stone unturned in ensuring his success next year. \nWe would recommend for you to meet {student_name}'s professors, so that you may better understand his needs, requirements and shortcomings. \nWe hope you understand the mental pressure {student_name} is going through, and we urge you not to reprimand him at present. His success is our duty, and we shall accomplish it by all means we can. \nSincerely, \n{college_dean}\nDean, {college_name} {college_city}"


        # Send an email to the student
        msg = MIMEText(email_body)
        msg['Subject'] = 'You Have Not Passed'
        msg['From'] = email_username
        msg['To'] = student_email

        server = smtplib.SMTP(email_server, email_port)
        server.starttls()
        server.login(email_username, email_password)
        server.sendmail(email_username, student_email, msg.as_string())
        server.quit()

        print(f'Email sent to {student_name} at {student_email}')
