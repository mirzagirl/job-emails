import smtplib
import pandas as pd
from email.message import EmailMessage
import os
import sys

try:
    user_limit = int(sys.argv[1])  # Get number input from command line
except (IndexError, ValueError):
    print("❌ Please provide a valid number. Example usage: python3 main.py 3")
    sys.exit(1)

#load files
data = pd.read_excel("modified_contacts.xlsx")

#email details
EMAIL = "mirzaalina169@gmail.com"
PWD = "you - password from my passwor in google security"

# Job role-specific email templates
TEMPLATES = {
    "frontend_pune": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am writing to express my interest in the <strong>Frontend Developer</strong>  at <strong>{company}</strong>. 
        With a strong foundation in React.js, HTML/CSS, and UI/UX design, I am eager to contribute to your team and further 
        develop my skills.</p>

        <p>Please find my resume attached for your review. I would love the opportunity to discuss how I can add value 
        to your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892 <br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "backend_pune": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am excited to apply for the <strong>Backend Developer</strong>  at <strong>{company}</strong>. 
        My experience with Node.js, Express, and database management (SQL & MongoDB) has prepared me to take on 
        backend development challenges.</p>

        <p>Enclosed is my resume for your review. I look forward to discussing how my skills align with the needs of your company.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "fullstack_pune": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am reaching out to apply for the <strong>Full Stack Developer</strong>  at <strong>{company}</strong>. 
        My expertise in both frontend (React.js) and backend (Node.js, Express) development allows me to contribute to 
        all aspects of web applications.</p>

        <p>Attached is my resume for your review. I look forward to the possibility of working with your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,
        "frontend_banglore": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am writing to express my interest in the <strong>Frontend Developer</strong>  at <strong>{company}</strong>. 
        With a strong foundation in React.js, HTML/CSS, and UI/UX design, I am eager to contribute to your team and further 
        develop my skills.</p>

        <p>Please find my resume attached for your review. I would love the opportunity to discuss how I can add value 
        to your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "backend_banglore": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am excited to apply for the <strong>Backend Developer</strong>  at <strong>{company}</strong>. 
        My experience with Node.js, Express, and database management (SQL & MongoDB) has prepared me to take on 
        backend development challenges.</p>

        <p>Enclosed is my resume for your review. I look forward to discussing how my skills align with the needs of your company.</p>
       
        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "fullstack_banglore": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am reaching out to apply for the <strong>Full Stack Developer</strong>  at <strong>{company}</strong>. 
        My expertise in both frontend (React.js) and backend (Node.js, Express) development allows me to contribute to 
        all aspects of web applications.</p>

        <p>Attached is my resume for your review. I look forward to the possibility of working with your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>
   
        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,
        "frontend_banglore_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am writing to express my interest in the <strong>Frontend Developer</strong>  at <strong>{company}</strong>. 
        With a strong foundation in React.js, HTML/CSS, and UI/UX design, I am eager to contribute to your team and further 
        develop my skills.</p>

        <p>Please find my resume attached for your review. I would love the opportunity to discuss how I can add value 
        to your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "backend_banglore_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am excited to apply for the <strong>Backend Developer</strong>  at <strong>{company}</strong>. 
        My experience with Node.js, Express, and database management (SQL & MongoDB) has prepared me to take on 
        backend development challenges.</p>

        <p>Enclosed is my resume for your review. I look forward to discussing how my skills align with the needs of your company.</p>
       
        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "fullstack_banglore_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am reaching out to apply for the <strong>Full Stack Developer</strong>  at <strong>{company}</strong>. 
        My expertise in both frontend (React.js) and backend (Node.js, Express) development allows me to contribute to 
        all aspects of web applications.</p>

        <p>Attached is my resume for your review. I look forward to the possibility of working with your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>
   
        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,
     "frontend_pune_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am writing to express my interest in the <strong>Frontend Developer</strong>  at <strong>{company}</strong>. 
        With a strong foundation in React.js, HTML/CSS, and UI/UX design, I am eager to contribute to your team and further 
        develop my skills.</p>

        <p>Please find my resume attached for your review. I would love the opportunity to discuss how I can add value 
        to your team.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892 <br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "backend_pune_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am excited to apply for the <strong>Backend Developer</strong>  at <strong>{company}</strong>. 
        My experience with Node.js, Express, and database management (SQL & MongoDB) has prepared me to take on 
        backend development challenges.</p>

        <p>Enclosed is my resume for your review. I look forward to discussing how my skills align with the needs of your company.</p>

        <p>Looking forward to your response.</p> <br> Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,

    "fullstack_pune_rf": """ 
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.5;">
        <p style="font-size: 18px;"><strong>Dear {hr_name},</strong></p>

        <p>I am reaching out to apply for the <strong>Full Stack Developer</strong>  at <strong>{company}</strong>. 
        My expertise in both frontend (React.js) and backend (Node.js, Express) development allows me to contribute to 
        all aspects of web applications.</p>

        <p>Attached is my resume for your review. I look forward to the possibility of working with your team.</p>

        <p>Looking forward to your response.</p> <br>  Job Link I'm Interested: <strong>{job_link}</strong>

        <p>Best regards,<br>
        <strong>Alina Mirza</strong><br>
        Contact: 9131292892<br>
        Email: alinamirza169@gmail.com</p>
    </body>
    </html>
    """,
}

# Resume file paths based on job roles
RESUMES = {
    "frontend_pune": "alinaM_pf.pdf",
    "backend_pune": "alinaM_pb.pdf",
    "fullstack_pune": "alinaM_pfs.pdf",
    "frontend_banglore": "alinaM_bf.pdf",
    "backend_banglore": "alinaM_bb.pdf",
    "fullstack_banglore": "alinaM_bfs.pdf",
    "frontend_pune_rf": "alinaM_pf.pdf",
    "backend_pune_rf": "alinaM_pb.pdf",
    "fullstack_pune_rf": "alinaM_pfs.pdf",
    "frontend_banglore_rf": "alinaM_bf.pdf",
    "backend_banglore_rf": "alinaM_bb.pdf",
    "fullstack_banglore_rf": "alinaM_bfs.pdf",
}

# Email sender function
def email_sender(to_mail, hr_name, company,job_title, job_role,job_link):
    # Validate job role
    if job_role not in TEMPLATES or job_role not in RESUMES:
        print(f"❌ Invalid job role: {job_role}. Skipping email to {to_mail}.")
        return

    # Get the appropriate template and resume
    body = TEMPLATES[job_role].format(hr_name=hr_name, company=company,job_link=job_link)
    resume_path = RESUMES[job_role]

    # Prepare email
    mail = EmailMessage()
    mail["From"] = EMAIL
    mail["To"] = to_mail
    mail["Subject"] = f"Application for {job_title.capitalize()} Developer at {company}"

    # Set HTML email content
    mail.set_content("This email requires an HTML-compatible email client.")  # Fallback text
    mail.add_alternative(body, subtype="html")  # HTML content

    # Attach Resume
    if os.path.exists(resume_path):
        with open(resume_path, "rb") as file:
            mail.add_attachment(file.read(), maintype="application", subtype="pdf", filename=os.path.basename(resume_path))
    else:
        print(f"⚠️ Resume not found: {resume_path}. Email sent without attachment.")

    # Send email
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL, PWD)
        server.send_message(mail)
        server.quit()
        print(f"✅ Email sent successfully to {to_mail} for {job_title} role.")
    except Exception as e:
        print(f"❌ Email not sent to {to_mail}. Error: {e}")

# Example usage

for index,row in data.iterrows():
    if(index<user_limit):
      continue
    email_sender(row["Email"],row["Name"],row["Company"],row["Jobrole"],row["Role"],row["Joblink"])
