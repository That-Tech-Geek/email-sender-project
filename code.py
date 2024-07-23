import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Create a Streamlit app
st.title("Student Result Checker")

# Upload CSV file
uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

if uploaded_file is not None:
    # Read the CSV file
    df = pd.read_csv(uploaded_file)

    # Get passing score from user
    passing_score = st.number_input("Enter passing score:", min_value=0, max_value=100)

    # Check which students have passed and which failed
    df["Result"] = df["Marks Scored"].apply(lambda x: "Passed" if x >= passing_score else "Failed")

    # Create a table with easily identifiable distinctions
    st.write("Student Results:")
    st.write(df[["Name", "Roll No.", "Email ID of Guardian", "Marks Scored", "Result"]])

    # Create separate tables for passed and failed students
    passed_students = df[df["Result"] == "Passed"]
    failed_students = df[df["Result"] == "Failed"]

    st.write("Passed Students:")
    st.write(passed_students[["Name", "Roll No.", "Email ID of Guardian", "Marks Scored"]])

    st.write("Failed Students:")
    st.write(failed_students[["Name", "Roll No.", "Email ID of Guardian", "Marks Scored"]])

    # Send emails to students
    st.write("Enter email ID and app password to send emails:")
    email_id = st.text_input("Email ID:")
    app_password = st.text_input("App Password:", type="password")

    if st.button("Send Emails"):
        # Create a SMTP server
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_id, app_password)

        # Send emails to students
        for index, row in df.iterrows():
            msg = MIMEMultipart()
            msg["From"] = email_id
            msg["To"] = row["Email ID of Guardian"]
            msg["Subject"] = "Student Result"

            if row["Result"] == "Passed":
                body = f"Dear {row['Name']}\'s Guardian,\n\n Congratulations! {row['Name']} has passed the exam.\n\nBest regards,\n[Your Name]"
            else:
                body = f"Dear {row['Name']}\'s Guardian,\n\n Sorry to inform that {row['Name']} has failed the exam.\n\nBest regards,\n[Your Name]"

            msg.attach(MIMEText(body, "plain"))
            server.send_message(msg)

        server.quit()
        st.write("Emails sent successfully!")
