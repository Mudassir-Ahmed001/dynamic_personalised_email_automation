

import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import re
from st_chat_message import message  # For the 'Email Content Generation' menu
import requests

# Set page configuration
st.set_page_config(page_title="Dynamic Email Automation", layout="centered")

# Importing the custom menu
from streamlit_option_menu import option_menu

# Function to validate email
def is_valid_email(email):
    return bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email))

# Header Menu
with st.sidebar:
    selected = option_menu(
        "Menu",
        ["Email Automation", "Email Content Generation", "How to Use"],
        icons=["envelope", "pencil", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )

# Menu: Email Automation
if selected == "Email Automation":
    st.title("Email Automation")
    st.markdown("Send personalized emails with dynamic attachments and content.")

    # Input for email and App Password
    email = st.text_input("Your Email", type="default", help="Enter your email address (e.g., Gmail).")
    app_password = st.text_input("App Password", type="password", help="Enter your email App Password.")

    # Validation for email
    if email and not is_valid_email(email):
        st.error("Invalid email address format. Please correct it.")

    # Upload recipient sheet
    uploaded_file = st.file_uploader("Upload Recipient Sheet (.csv or .xlsx)", type=["csv", "xlsx"])

    data = None  # Ensure data is always defined
    invalid_emails = pd.DataFrame()  # Initialize empty DataFrame for invalid emails

    if uploaded_file:
        try:
            # Read file and validate structure
            file_ext = uploaded_file.name.split(".")[-1]
            if file_ext == "csv":
                data = pd.read_csv(uploaded_file)
            else:
                data = pd.read_excel(uploaded_file)

            data.columns = map(str.lower, data.columns)
            data.columns = data.columns.str.strip()

            if "name" not in data.columns or "email" not in data.columns:
                st.error("The uploaded sheet must contain columns named 'Name' and 'Email'.")
                data = None  # Invalidate data
            else:
                # Validate emails
                data["is_valid_email"] = data["email"].apply(lambda x: bool(is_valid_email(x)))
                invalid_emails = data[data["is_valid_email"] == False]  # Correct filtering

                data = data[data["is_valid_email"]]

                if not data.empty:
                    personalization_vars = ", ".join([f"{{{col}}}" for col in data.columns if col not in ["is_valid_email"]])  
                    st.markdown(f"### The personalization variables available are: {personalization_vars}")

                    st.markdown("### Sample Data:")
                    st.dataframe(data.head())

                    if not invalid_emails.empty:
                        st.warning(f"{len(invalid_emails)} invalid email(s) found and excluded:")
                        st.dataframe(invalid_emails)
                else:
                    st.error("No valid email addresses found in the sheet. Please upload a valid file.")

        except Exception as e:
            st.error(f"An error occurred while processing the file: {str(e)}")

    # Upload certificates
    cert_files = st.file_uploader("Upload Certificates (PDFs)", type=["pdf", "png", "jpg"], accept_multiple_files=True)

    # Upload constant attachments
    constant_attachments = st.file_uploader("Upload Constant Attachments (Optional)", type=["pdf", "docx", "png", "jpg"], accept_multiple_files=True)

    # Input for email subject and content
    email_subject = st.text_input("Email Subject", help="Subject of the email. Use placeholders like {name}, {email}.")
    email_content = st.text_area("Email Content (HTML supported)", height=200, help="Write your email content with placeholders like {name}.")

    # Send Button (always visible)
    if st.button("Send Emails"):
        if not email or not app_password:
            st.error("Please provide your email and App Password before sending emails.")
        elif data is None or data.empty:
            st.error("No valid recipient data found. Please upload a valid sheet.")
        else:
            try:
                matched_files = {}
                for cert in cert_files:
                    cert_name = os.path.splitext(cert.name)[0].strip().lower().replace(" ", "")
                    matched_files[cert_name] = cert

                progress = st.progress(0)
                total_recipients = len(data)
                emails_sent = 0

                for index, row in data.iterrows():
                    name = str(row["name"]).strip().lower().replace(" ", "")
                    recipient_email = row["email"]

                    if name in matched_files:
                        message = MIMEMultipart()
                        message["From"] = email
                        message["To"] = recipient_email
                        message["Subject"] = email_subject.format(**row)

                        body = email_content.replace("\n", "<br>").format(**row)  # Preserve new lines
                        message.attach(MIMEText(body, "html"))


                        attachment = matched_files[name]
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={attachment.name}")
                        message.attach(part)

                        for const_file in constant_attachments:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(const_file.read())
                            encoders.encode_base64(part)
                            part.add_header("Content-Disposition", f"attachment; filename={const_file.name}")
                            message.attach(part)

                        with smtplib.SMTP("smtp.gmail.com", 587) as server:
                            server.starttls()
                            server.login(email, app_password)
                            server.send_message(message)

                        emails_sent += 1
                        st.success(f"Email sent to {recipient_email}")
                    else:
                        st.warning(f"No matching certificate for {name}. Skipping email.")

                    progress.progress((index + 1) / total_recipients)

                progress.progress(1.0)
                st.success(f"Emails sent: {emails_sent}/{total_recipients}")

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

# Menu: Email Content Generation
elif selected == "Email Content Generation":
    st.title("Email Content Generation")
    st.markdown("Generate professional, formal emails or enhance your existing drafts with 'Llama 3.3 70b'")

    # Let the user enter their GROQ API Key instead of hardcoding it
    groq_api_key = "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"

    # User input for email content
    user_input = st.text_area("Enter Context or Email Draft", height=200)

    if st.button("Generate Content"):
        if not groq_api_key or not user_input:
            st.error("Input is required.")
        else:
            headers = {
                "Authorization": f"Bearer {groq_api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "llama-3.3-70b-versatile",  # Correct model name
                "messages": [{"role": "user", "content": user_input}],
                "temperature": 0.3
            }

            try:
                response = requests.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers=headers,
                    json=payload
                )

                if response.status_code == 200:
                    response_text = response.json()["choices"][0]["message"]["content"]
                    st.markdown("### Generated Email Content:")
                    st.write(response_text)

                else:
                    st.error(f"Failed to generate content: {response.status_code} - {response.json()}")
            except Exception as e:
                st.error(f"An error occurred while connecting to the GROQ API: {str(e)}")

elif selected == "How to Use":
    st.title("How to Use the App")
    st.markdown("""
    ## 📌 Features of This App:
    - **Email Automation**: Upload recipient data, certificates, and constant attachments to send personalized emails.
    - **Email Content Generation**: Use AI to generate or enhance email content.
    
    ---

    ## 📧 **How to Use Email Automation**
    
    1️⃣ **Enter Your Email and App Password**  
       - Go to **Email Automation**.  
       - Enter your **Gmail address** and **App Password** (see below to get it).
    
    2️⃣ **Upload the Recipient Sheet**  
       - The file must be in **CSV** or **Excel** format.  
       - It must contain at least two columns:  
         - **`name`** → The recipient's name  
         - **`email`** → The recipient’s email address  
       - The app will automatically **detect invalid emails** and remove them.  

    3️⃣ **Upload Certificates**  
       - The certificates **must be named exactly** as in the `name` column of the recipient sheet.  
       - **Example:**  
         - If the sheet has `John Doe`, the file name should be **"John Doe.pdf"** (case and space insensitive).  
         - `JOHN DOE.pdf`, `john_doe.pdf`, or `JohnDoe.pdf` will all match correctly.  

    4️⃣ **Upload Constant Attachments (Optional)**  
       - You can upload any additional **documents, PDFs, or images** that should be attached to all emails.  

    5️⃣ **Enter Email Subject & Content**  
       - Use placeholders to personalize emails:  
         - `{name}` → Will be replaced with recipient's name  
         - `{email}` → Will be replaced with recipient’s email  
       - **Example:**  
         ```
         Hello {name},
         
         Here is your certificate of completion.
         
         Regards,
         Your Team
         ```
       - **New lines will be preserved** in the email.  

    6️⃣ **Click "Send Emails"**  
       - The app will send personalized emails to all recipients with their respective attachments.  
       - A progress bar will show the email-sending process.  

    ---

    ## 🔑 **How to Get Your Gmail App Password**
    Gmail requires **App Passwords** for third-party apps like this one.  
    Follow these steps to generate it:

    1. Open **Google Account Security Settings**:  
       👉 [https://myaccount.google.com/security](https://myaccount.google.com/security)  

    2. Scroll to **"Signing in to Google"** and enable **2-Step Verification** (if not already enabled).  

    3. Click **"App Passwords"** (If you don't see this option, enable **2-Step Verification** first).  

    4. Under **"Select App"**, choose **Mail**.  

    5. Under **"Select Device"**, choose **Other (Custom name)** and enter any name (e.g., "Email Automation App").  

    6. Click **"Generate"** → You will see a **16-character password** (e.g., `abcd efgh ijkl mnop`).  

    7. **Copy this password** and paste it into the **App Password** field in the app.  

    ⚠️ **DO NOT share this password.** It is used only for email sending via this app.

    ---

    ## ✨ **How to Use Email Content Generation**
    
    1️⃣ **Enter Context or an Email Draft**  
       - Example 1: `"Generate a formal email inviting employees to an annual meeting."`  
       - Example 2: `"Improve this email: Hi John, please complete your profile update."` 

    2️⃣  **Click "Generate Content"**  
       - The AI will generate a professional email based on your input.  

    ---

    ## ❗ Important Notes
    - ✅ **Ensure recipient names in the sheet match certificate file names exactly.**  
    - ✅ **Use `{name}` and `{email}` for personalization.**  
    - ✅ **Make sure the email content is properly formatted to maintain new lines.**  
    - ✅ **App Password is required for sending emails via Gmail.**  

    🚀 You're all set to send personalized emails!  
    """)

