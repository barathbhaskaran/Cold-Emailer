import yagmail
import schedule
import time
import pandas as pd
import logging

# Set up logging for error handling
logging.basicConfig(filename="email_bot.log", level=logging.ERROR,
                    format="%(asctime)s %(levelname)s %(message)s")


# Initialize the SMTP client and send the email
def send_email(recipient, subject, content, attachment):
    try:
        yag = yagmail.SMTP('', '')

        # Send the email with an attachment
        yag.send(
            to=recipient,
            subject=subject,
            contents=content,
            attachments=attachment  # Attaching the resume PDF
        )
        print(f"Email sent to {recipient} with attachment")

    except Exception as e:
        logging.error(f"Failed to send email to {recipient}: {e}")
        print(f"Error: Failed to send email to {recipient}")


# Load recipients from a CSV file
def load_recipients_from_csv(file_path):
    try:
        data = pd.read_excel(file_path)
        print(data.head())
        return data
    except Exception as e:
        logging.error(f"Failed to load CSV file: {e}")
        return pd.DataFrame()  # Return an empty DataFrame if there's an error


# Create dynamic content for each recipient using their name
def create_email_content(name):
    return (f"Hi {name},\n\nI hope this message finds you well. My name is Barath Bhaskaran, and I recently completed my Master’s in Software Engineering at the University of Maryland. With a strong foundation in backend development, cloud computing, and CI/CD automation, I have been fortunate to work on projects that significantly improved user engagement and operational efficiency at the organizations I've served."
            "\n In my most recent position as a software engineer at the University of Maryland, I engineered backend website operations that led to a 90% increase in website customization and user engagement. Prior to this, during my tenure at Infosys Ltd (Caterpillar), I developed an innovative chatbot that optimized user request handling, resulting in an 80% increase in task throughput. I also spearheaded the CI/CD automation using Azure DevOps, reducing manual intervention by 80%."
            "\n I am highly skilled in Java, Spring Boot, and various web development frameworks such as React and Angular. My technical arsenal also includes proficiency in Docker, AWS, and REST API development, among others."
            "\nThank you for considering my application. I look forward to the possibility of contributing to your team."
            "\n\nBest regards,\n Barath Bhaskaran \n[https://www.linkedin.com/in/barath-bhaskaran/] | \n[https://github.com/barathbhaskaran] | \n [https://barathbhaskaran.github.io/]")


# Function to send emails to all recipients from the CSV
def send_emails_to_all(file_path, attachment_path):
    recipients_data = load_recipients_from_csv(file_path)

    if recipients_data.empty:
        print("No recipients to send emails to. Check the CSV file.")
        return

    # Loop through each row in the CSV
    for index, row in recipients_data.iterrows():
        recipient = row['email']  # Fetch email
        name = row['name']  # Fetch name

        # Create dynamic subject and content
        subject = "Regarding Software Engineer Opportunity"
        content = create_email_content(name)

        send_email(recipient, subject, content, attachment_path)


# # Schedule the email to be sent at a specific time
# def schedule_email_sending(file_path, attachment_path):
#     schedule.every().day.at("10:00").do(send_emails_to_all, file_path=file_path, attachment_path=attachment_path)


# Keep the bot running
def run_bot(file_path, attachment_path):
    send_emails_to_all(file_path, attachment_path)

    while True:
        schedule.run_pending()
        time.sleep(60)  # Wait 1 minute before checking the schedule again


if __name__ == "__main__":
    # Set the path to your CSV file
    csv_file_path = r"E:\pythonProject\input.xlsx"

    # Set the path to your resume PDF file
    resume_pdf_path = r"E:\pythonProject\Barath_Bhaskaran_Resume.pdf"  # Update this to the correct file path

    # Start the bot
    run_bot(csv_file_path, resume_pdf_path)
