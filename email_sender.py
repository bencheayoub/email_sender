import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
from datetime import date
import os

# Configuration
GMAIL = "your_email@gmail.com"
APP_PASSWORD = "your_app_password"

# Email templates (to be filled by users)
content = """
Your HTML email content will go here.

Available placeholders:
{{FIRST_NAME}} - User's first name
{{BUTTON_AREA}} - Button HTML section
"""

subject_accepted = "Your Accepted Email Subject"
subject_rejected = "Your Rejected Email Subject"

# Button HTML Structure
BUTTON_HTML_STRUCTURE = """
<div class="button-area">
    <a href="{LINK}" class="action-button">
        <span class="button-icon">{EMOJI}</span>
        {BUTTON_TEXT}
    </a>
</div>
"""

# Main Email Template
EMAIL_HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{SUBJECT}</title>
    <style>
        /* Your CSS styles will go here */
    </style>
</head>
<body>
    {CONTENT}
</body>
</html>
"""


def send_message(
        from_email,
        to_email,
        first_name,
        content,
        subject,
        link,
        button_text="Click Here",
        emoji="✨"
):
    """
    Send an email with HTML content and embedded images.
    
    Args:
        from_email: Sender's email address
        to_email: Recipient's email address
        first_name: Recipient's first name
        content: Email content HTML with placeholders
        subject: Email subject
        link: Button link URL
        button_text: Text for the call-to-action button
        emoji: Emoji for the button
    
    Returns:
        None
    """
    # 1. Render the button HTML
    button_html_rendered = BUTTON_HTML_STRUCTURE.format(
        LINK=link,
        BUTTON_TEXT=button_text,
        EMOJI=emoji
    )

    # 2. Replace placeholders in the content template
    final_content = content.replace("{{FIRST_NAME}}", first_name).replace("{{BUTTON_AREA}}", button_html_rendered)

    # 3. Insert final content into the main HTML template
    html = EMAIL_HTML_TEMPLATE.format(
        SUBJECT=subject,
        CONTENT=final_content
    )

    try:
        msg = MIMEMultipart("related")
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Subject"] = subject

        # Attach HTML content
        msg_alternative = MIMEMultipart("alternative")
        msg.attach(msg_alternative)
        msg_alternative.attach(MIMEText(html, "html"))

        # Attach logo image if exists
        try:
            with open("logo.png", "rb") as img_file:
                img = MIMEImage(img_file.read())
                img.add_header("Content-ID", "<logo_image>")
                img.add_header("Content-Disposition", "inline", filename="logo.png")
                msg.attach(img)
        except FileNotFoundError:
            print("Warning: logo.png not found. Email will be sent without logo.")

        # Attach social media icons
        social_icons = {
            "facebook_image": "facebook.png",
            "instagram_image": "instagram.png",
            "linkedin_image": "linkedin.png",
            "github_image": "github.png"
        }

        for content_id, filename in social_icons.items():
            try:
                with open(filename, "rb") as img_file:
                    img = MIMEImage(img_file.read())
                    img.add_header("Content-ID", f"<{content_id}>")
                    img.add_header("Content-Disposition", "inline", filename=filename)
                    msg.attach(img)
            except FileNotFoundError:
                print(f"Warning: {filename} not found. Social icon will be missing.")

        # Send email
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(from_email, APP_PASSWORD)
            server.sendmail(from_email, to_email, msg.as_string())

        print(f"Email sent successfully to {to_email}")

    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def read_data_from_excel(file_path="applicants.xlsx"):
    """
    Read applicant data from Excel file.
    
    Args:
        file_path: Path to Excel file
        
    Returns:
        List of dictionaries containing user data
    """
    try:
        df = pd.read_excel(file_path)
        users = []

        for _, row in df.iterrows():
            # Extract user data (modify column names as needed)
            first_name = str(row['first_name']).strip() if 'first_name' in row else ""
            
            # Determine result based on status column
            status = str(row.get('status', '')).strip().lower() if 'status' in row else ""
            result = "accepted" if status == "accepted" else "rejected"

            user_data = {
                "email": str(row['email']).strip() if 'email' in row else "",
                "first_name": first_name,
                "result": result,
                "discord_invite_link": str(row.get('discord_link', 'https://discord.gg/your-invite-link')).strip()
            }
            users.append(user_data)

        return users
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []


def send_emails_to_all_users():
    """
    Send emails to all applicants based on their application status.
    """
    users = read_data_from_excel()

    if not users:
        print("No users found in the Excel file.")
        return

    for user in users:
        if user["result"] == "accepted":
            send_message(
                GMAIL,
                user["email"],
                user["first_name"],
                content,
                subject_accepted,
                user.get("discord_invite_link", "https://example.com"),
                button_text="Enter the Portal",
                emoji="🚀"
            )
        else:
            send_message(
                GMAIL,
                user["email"],
                user["first_name"],
                content,
                subject_rejected,
                "https://example.com/community",  # Default community link
                button_text="Join Community",
                emoji="🦉"
            )


# Main execution
if __name__ == "__main__":
    print("Email Sender Script")
    print("==================")
    print("Please configure the following before running:")
    print("1. Set GMAIL and APP_PASSWORD variables")
    print("2. Update email templates (content, subject_accepted, subject_rejected)")
    print("3. Update EMAIL_HTML_TEMPLATE with your CSS and HTML structure")
    print("4. Prepare your Excel file with columns: email, first_name, status")
    print("5. Add your images (logo.png, social media icons)")
    print("\nRun send_emails_to_all_users() to start sending emails.")
    
    # Uncomment the line below to run automatically
    # send_emails_to_all_users()
