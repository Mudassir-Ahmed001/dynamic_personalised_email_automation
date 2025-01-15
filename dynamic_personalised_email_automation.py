import streamlit as st
import re
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
from time import sleep
from typing import List, Optional, Dict
import pandas as pd


class DebugLogger:
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode

    def debug(self, message: str):
        if self.debug_mode:
            print(f"[DEBUG] {message}")


class EmailValidator:
    @staticmethod
    def is_valid_email(email: str) -> bool:
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))


class EmailAutomation:
    def __init__(self, debug_mode: bool = False):
        self.debug_logger = DebugLogger(debug_mode)
        self.logger = self._setup_logging()
        self.smtp_server = None

    def _setup_logging(self) -> logging.Logger:
        logger = logging.getLogger('CertificateEmailer')
        logger.setLevel(logging.INFO)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_handler = logging.FileHandler(f'email_log_{timestamp}.txt')
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        return logger

    def setup_smtp(self, email: str, password: str):
        """Setup SMTP connection with Gmail."""
        self.debug_logger.debug("Setting up SMTP connection")
        try:
            self.smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
            self.smtp_server.starttls()
            self.smtp_server.login(email, password)
            self.logger.info("SMTP connection established successfully")
        except Exception as e:
            self.logger.error(f"SMTP setup failed: {str(e)}")
            raise

    def match_certificate(self, certificates: List[tuple], recipient_name: str) -> Optional[tuple]:
        """Match the certificate file with the recipient name."""
        recipient_name = recipient_name.lower().strip()
        for cert_name, cert_content in certificates:
            if recipient_name in cert_name.lower():
                return cert_name, cert_content
        return None

    def send_emails(self, sender: str, recipients_df: pd.DataFrame, subject_template: str,
                    content_template: str, certificates: List[tuple],
                    additional_attachments: Optional[List[tuple]] = None):
        """Send personalized emails with certificates to all recipients."""
        self.debug_logger.debug("Starting email sending process")
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, row in recipients_df.iterrows():
            recipient_data = row.to_dict()
            recipient_name = recipient_data.get('name', '').strip()
            recipient_email = recipient_data.get('email', '').strip()

            # Match the certificate for the current recipient
            certificate = self.match_certificate(certificates, recipient_name)

            retries = 3
            while retries > 0:
                try:
                    message = self.create_message(sender, recipient_data, subject_template,
                                                  content_template, certificate,
                                                  additional_attachments)
                    self.smtp_server.send_message(message)
                    self.logger.info(f"Email sent successfully to {recipient_email}")
                    status_text.text(f"Sent email to: {recipient_email}")
                    progress_bar.progress((idx + 1) / len(recipients_df))
                    sleep(1)
                    break
                except Exception as e:
                    retries -= 1
                    self.logger.error(f"Failed to send email to {recipient_email}: {str(e)}")
                    if retries > 0:
                        status_text.text(f"Retrying email... ({retries} attempts remaining)")
                        sleep(2)
                    else:
                        status_text.text(f"Failed to send email after maximum retries")

        status_text.text("Email campaign completed!")

    def create_message(self, sender: str, recipient_data: Dict, subject_template: str,
                       content_template: str, certificate: Optional[tuple],
                       additional_attachments: Optional[List[tuple]] = None) -> MIMEMultipart:
        """Create personalized email message."""
        recipient_email = recipient_data.get('email', '')
        personalized_subject = subject_template.format(**recipient_data)
        personalized_content = content_template.format(**recipient_data)

        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = recipient_email
        message['Subject'] = personalized_subject
        message.attach(MIMEText(personalized_content, 'html', 'utf-8'))

        # Attach the matched certificate
        if certificate:
            cert_name, cert_content = certificate
            file_ext = cert_name.lower().split('.')[-1]
            if file_ext == 'pdf':
                attachment = MIMEApplication(cert_content, _subtype="pdf")
            elif file_ext in ['jpg', 'jpeg', 'png']:
                attachment = MIMEImage(cert_content)
            else:
                raise ValueError(f"Unsupported certificate format: {file_ext}")
            attachment.add_header('Content-Disposition', 'attachment', filename=cert_name)
            message.attach(attachment)

        # Attach additional files if provided
        if additional_attachments:
            for file_name, file_content in additional_attachments:
                try:
                    file_ext = file_name.lower().split('.')[-1]
                    if file_ext == 'pdf':
                        attachment = MIMEApplication(file_content, _subtype="pdf")
                    else:
                        attachment = MIMEImage(file_content)
                    attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
                    message.attach(attachment)
                except Exception as e:
                    self.logger.error(f"Error attaching file {file_name}: {str(e)}")
                    raise

        return message


def display_how_to_use():
    """Display 'How to Use' popup with updated content."""
    st.markdown("""
    ## 📖 How to Use the dynamic personalized email automation system

    ### Supported File Types
    - Currently, **only `.xlsx` files are supported**.

    ### Instructions
    1. Enter the email and respective email APP password.
        How to get email app password:
        
            - After enabling 2-Step Verification, go back to the Security page
            
            - Scroll down to find "App passwords" (it only appears after 2-Step Verification is enabled)
            
            - You might need to sign in again
            
            - Click on "Select app" dropdown and choose "Mail"
            
            - For "Select device" choose "Other (Custom name)"
            
            - Give it a name like "Python Email Script"
            
            - Click "Generate"
            
            - Google will display a 16-character password (like: xxxx xxxx xxxx xxxx)
            
    2. Prepare an `.xlsx` file with the following required columns:
       - `name`: Recipient's name
       - `email`: Recipient's email address
    3. Upload the `.xlsx` file using the uploader.
    4. Personalize your email content using `{variable_name}` placeholders.
    5. Attach certificates and any additional files if needed.
    6. Enter your email credentials and send the campaign.
    """)


def main():
    st.title("bhayankar aangaar gir gir bolke bhejo emails ab khali")

    # Display "How to Use" Button on Main Page
    if st.button("📖 How to Use"):
        display_how_to_use()

    # Sender credentials
    sender_email = st.text_input("Sender Email")
    st.info("Not the email password! APP PASSWORD for the respective email(read the 'How to Use' for more clarity)")
    sender_password = st.text_input("Sender APP Password", type="password")

    # File uploader for `.xlsx` files
    st.write("Upload recipient data file (.xlsx only)")
    recipient_file = st.file_uploader("Upload File", type=['xlsx'])

    recipients_df = None  # Initialize recipients_df

    if recipient_file is not None:
        try:
            # Read Excel file
            recipients_df = pd.read_excel(recipient_file, engine='openpyxl')
            recipients_df.columns = [col.lower().strip() for col in recipients_df.columns]

            if 'name' not in recipients_df.columns or 'email' not in recipients_df.columns:
                st.error("File must contain 'name' and 'email' columns.")
                return

            # Validate emails
            recipients_df = recipients_df[recipients_df['email'].apply(EmailValidator.is_valid_email)]

            st.write("Data Preview (first 5 rows):")
            st.write(recipients_df.head())

            st.write("Available variables for personalization:")
            for col in recipients_df.columns:
                st.code(f"{{{{{col}}}}}")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

    # Certificate uploads
    st.write("Upload certificates (names should match recipient names in the .xlsx file)")
    certificate_files = st.file_uploader(
        "Upload certificates",
        type=['pdf', 'jpg', 'jpeg', 'png'],
        accept_multiple_files=True
    )

    # Additional attachments
    additional_attachments = st.file_uploader(
        "Upload additional attachments (optional)",
        type=['jpg', 'jpeg', 'png', 'pdf'],
        accept_multiple_files=True
    )

    # Email content
    st.subheader("Email Content")
    st.write("Use {variable_name} syntax for personalization")
    subject_template = st.text_input("Email Subject Template")
    content_template = st.text_area("Email Content Template (HTML supported)")
    content_template = content_template.replace("\xa0", "")

    if st.button("Send Emails"):
        try:
            if recipients_df is None:
                st.error("Please upload a valid .xlsx recipient data file.")
                return

            if not certificate_files:
                st.error("Please upload certificate files.")
                return

            if not subject_template or not content_template:
                st.error("Please fill in both subject and content templates.")
                return

            automation = EmailAutomation(debug_mode=False)

            # Process certificates
            certificates = [(file.name, file.read()) for file in certificate_files]

            # Process additional attachments
            attachments = []
            if additional_attachments:
                for file in additional_attachments:
                    attachments.append((file.name, file.read()))

            # Setup SMTP
            automation.setup_smtp(sender_email, sender_password)

            # Send emails
            automation.send_emails(
                sender_email,
                recipients_df,
                subject_template,
                content_template,
                certificates,
                attachments
            )

            st.success("Email campaign completed successfully!")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            logging.error(f"Critical error: {str(e)}")
        finally:
            if automation.smtp_server:
                automation.smtp_server.quit()


if __name__ == "__main__":
    main()
