# streamlit_app.py

import streamlit as st
import pandas as pd
import smtplib
import ssl
import re

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# ---------------- PAGE CONFIG ---------------- #

st.set_page_config(
    page_title="Bulk Email Sender",
    layout="centered"
)

# ---------------- CUSTOM CSS ---------------- #

st.markdown("""
<style>

    .stApp {
        background-color: #1f1f1f;
        color: white;
    }

    .main-title {
        text-align: center;
        color: white;
        font-size: 26px;
        font-weight: bold;
        margin-bottom: 25px;
    }

    .block-container {
        padding-top: 2rem;
    }

    /* INPUT FIELDS */

    .stTextInput input {
        background-color: #2b2b2b !important;
        color: white !important;
        border: 1px solid #555 !important;
    }

    .stTextArea textarea {
        background-color: #2b2b2b !important;
        color: white !important;
        border: 1px solid #555 !important;
        font-weight: bold !important;
    }

    /* PLACEHOLDER TEXT */

    input::placeholder,
    textarea::placeholder {
        color: #cfcfcf !important;
        opacity: 1 !important;
    }

    /* FILE UPLOADER */

    .stFileUploader {
        background-color: #2b2b2b !important;
        padding: 10px;
        border-radius: 8px;
        border: 1px solid #555;
        color: white !important;
    }

    section[data-testid="stFileUploaderDropzone"] {
        background-color: #2b2b2b !important;
        color: white !important;
        border: 1px dashed #666 !important;
    }

    section[data-testid="stFileUploaderDropzone"] * {
        color: white !important;
    }

    /* LABELS */

    label {
        color: white !important;
        font-weight: 500;
    }

    /* BUTTON */

    .stButton button {
        width: 100%;
        background-color: #d32f2f;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 50px;
        border: none;
    }

    .stButton button:hover {
        background-color: #b71c1c;
        color: white;
    }

</style>
""", unsafe_allow_html=True)

# ---------------- TITLE ---------------- #

st.markdown(
    '<div class="main-title">Bulk Email Sender Application</div>',
    unsafe_allow_html=True
)

# ---------------- INPUTS ---------------- #

col1, col2 = st.columns(2)

with col1:
    sender_email = st.text_input("Sender Email")

with col2:
    app_password = st.text_input(
        "Email App Password",
        type="password"
    )

uploaded_file = st.file_uploader(
    "Receivers Email and Name (Upload Excel File)",
    type=["xlsx", "xls"]
)

cc_emails = st.text_area(
    "Cc: (emails separated by comma)",
    placeholder="xyz@email.com, abc@email.com"
)

bcc_emails = st.text_area(
    "Bcc: (Optional)",
    placeholder="xyz@email.com, abc@email.com"
)

subject = st.text_input("Subject")

# ---------------- RICH TEXT EMAIL EDITOR ---------------- #

st.markdown("""
<link href="https://cdn.quilljs.com/1.3.6/quill.snow.css" rel="stylesheet">
<script src="https://cdn.quilljs.com/1.3.6/quill.js"></script>

<style>

    /* Toolbar */

    .ql-toolbar.ql-snow {
        background-color: #2b2b2b !important;
        border: 1px solid #555 !important;
        border-radius: 8px 8px 0px 0px !important;
    }

    /* Toolbar Icons */

    .ql-toolbar button svg {
        filter: invert(1);
    }

    .ql-picker {
        color: white !important;
    }

    .ql-picker-label {
        color: white !important;
    }

    /* Editor Container */

    .ql-container.ql-snow {
        background-color: white !important;
        border: 1px solid #555 !important;
        border-top: none !important;
        border-radius: 0px 0px 8px 8px !important;
    }

    /* Editable Area */

    .ql-editor {
        min-height: 400px !important;

        background-color: white !important;
        color: black !important;

        font-family: Arial, sans-serif !important;
        font-size: 14px !important;
        line-height: 1.6 !important;

        caret-color: black !important;

        overflow-y: auto !important;
    }

    /* Lists */

    .ql-editor ul,
    .ql-editor ol {
        padding-left: 1.5rem !important;
    }

    /* Placeholder */

    .ql-editor.ql-blank::before {
        color: #888 !important;
        font-style: normal !important;
    }

</style>
""", unsafe_allow_html=True)

st.markdown("### Write Email Body Here")

default_email_body = """
<p>Hi [Owner Name],</p>

<p>
We're launching Dotpe Horizon, an AI-powered business intelligence digest,
built entirely from your Rista data to help grow your revenue.
</p>

<br>

<p>Every week, on WhatsApp, you'll get:</p>

<ul>
    <li>What's actually driving (or dragging) your revenue</li>
    <li>Where your orders, AOV, customers and margins moved - and why</li>
    <li>Actions to address before next week</li>
</ul>

<p>
It's private. It's yours. No benchmarks, no comparisons - just your numbers.
To start receiving it, simply reply to this email with "YES".
Your data is never shared with anyone outside Horizon.
</p>

<br>

<p>Team Dotpe Horizon</p>
"""

email_body = st.components.v1.html(
    f'''
    <div id="toolbar">

        <select class="ql-font"></select>
        <select class="ql-size"></select>

        <button class="ql-bold"></button>
        <button class="ql-italic"></button>
        <button class="ql-underline"></button>

        <button class="ql-list" value="ordered"></button>
        <button class="ql-list" value="bullet"></button>

        <button class="ql-link"></button>

        <button class="ql-clean"></button>

    </div>

    <div id="editor">
        {default_email_body}
    </div>

    <script>

        var quill = new Quill('#editor', {{

            modules: {{
                toolbar: '#toolbar'
            }},

            theme: 'snow'
        }});

        // Make editor editable properly

        document.querySelector('.ql-editor').setAttribute(
            'contenteditable',
            'true'
        );

    </script>
    ''',
    height=500,
    scrolling=True
)
# ---------------- SIGNATURE ---------------- #

signature = st.text_area(
    "Paste Email Signature Here",
    height=150,
    placeholder="""Regards,

Shashank Shandilya
Executive-Data Analytics
M:+918860844270"""
)

# ---------------- LOGO ---------------- #

logo_path = "dotpe_logo.png"

st.image(logo_path, width=180)

# ---------------- FUNCTIONS ---------------- #

def extract_greeting(text):

    greetings = [
        "Hi",
        "Hello",
        "Dear",
        "Good Morning",
        "Good Afternoon",
        "Good Evening",
        "Respected"
    ]

    for greeting in greetings:

        pattern = rf"\\b{greeting}\\b"

        if re.search(pattern, text, re.IGNORECASE):
            return greeting

    return None


def personalize_email(body, full_name):

    first_name = full_name.strip().split()[0]

    greeting = extract_greeting(body)

    if greeting:

        pattern = rf"{greeting}"

        replacement = f"{greeting} {first_name},"

        updated_body = re.sub(
            pattern,
            replacement,
            body,
            count=1,
            flags=re.IGNORECASE
        )

        return updated_body

    return body


def format_signature(signature_text):

    lines = signature_text.split("\\n")

    formatted_lines = []

    for line in lines:

        if line.strip():

            formatted_lines.append(
                f"<b>{line}</b>"
            )

        else:

            formatted_lines.append("<br>")

    return "<br>".join(formatted_lines)


def send_bulk_emails(
    sender_email,
    app_password,
    df,
    cc_emails,
    bcc_emails,
    subject,
    email_body,
    signature
):

    smtp_server = "smtp.gmail.com"
    port = 465

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(
        smtp_server,
        port,
        context=context
    ) as server:

        server.login(
            sender_email,
            app_password
        )

        for index, row in df.iterrows():

            receiver_email = str(
                row["Email"]
            ).strip()

            receiver_name = str(
                row["Name"]
            ).strip()

            personalized_body = personalize_email(
                email_body,
                receiver_name
            )

            formatted_signature = format_signature(
                signature
            )

            final_body = f"""
            <html>

            <body style="background-color:white;color:black;font-family:Arial;">

                <div style="font-size:14px;line-height:1.6;">

                    {personalized_body}

                    <br><br>

                    {formatted_signature}

                    <br><br>

                    <img src="cid:dotpelogo" width="180">

                </div>

            </body>

            </html>
            """

            msg = MIMEMultipart("related")

            msg["From"] = sender_email
            msg["To"] = receiver_email
            msg["Subject"] = subject

            if cc_emails.strip():
                msg["Cc"] = cc_emails

            recipients = [receiver_email]

            if cc_emails.strip():

                recipients += [
                    email.strip()
                    for email in cc_emails.split(",")
                ]

            if bcc_emails.strip():

                recipients += [
                    email.strip()
                    for email in bcc_emails.split(",")
                ]

            html_part = MIMEText(
                final_body,
                "html"
            )

            msg.attach(html_part)

            # Attach Logo

            with open(logo_path, "rb") as img:

                mime_img = MIMEImage(img.read())

                mime_img.add_header(
                    "Content-ID",
                    "<dotpelogo>"
                )

                msg.attach(mime_img)

            server.sendmail(
                sender_email,
                recipients,
                msg.as_string()
            )

# ---------------- VALIDATIONS ---------------- #

mandatory_fields = (
    sender_email,
    app_password,
    uploaded_file,
    cc_emails,
    subject,
    signature
)

# ---------------- SEND BUTTON ---------------- #

if st.button("Send Emails"):

    if not all(mandatory_fields):

        st.error(
            "Please fill all mandatory fields."
        )

    else:

        try:

            df = pd.read_excel(
                uploaded_file
            )

            required_columns = [
                "Email",
                "Name"
            ]

            if not all(
                col in df.columns
                for col in required_columns
            ):

                st.error(
                    "Excel file must contain Email and Name columns."
                )

            else:

                send_bulk_emails(
                    sender_email,
                    app_password,
                    df,
                    cc_emails,
                    bcc_emails,
                    subject,
                    email_body,
                    signature
                )

                st.success(
                    "Emails sent successfully!"
                )

        except Exception as e:

            st.error(f"Error: {str(e)}")
