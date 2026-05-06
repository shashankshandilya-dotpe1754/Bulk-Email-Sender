# streamlit_app.py

import streamlit as st
import pandas as pd
import smtplib
import ssl
import re
import mimetypes

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

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

    .stTextInput input {
        background-color: #2b2b2b !important;
        color: white !important;
        border: 1px solid #555 !important;
    }

    .stTextArea textarea {
        background-color: #2b2b2b !important;
        color: white !important;
        border: 1px solid #555 !important;
    }

    input::placeholder,
    textarea::placeholder {
        color: #cfcfcf !important;
        opacity: 1 !important;
    }

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

    label {
        color: white !important;
        font-weight: 500;
    }

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

# ---------------- QUILL EDITOR ---------------- #

email_body = st.components.v1.html(
    """
    
    <link href="https://cdn.quilljs.com/1.3.6/quill.snow.css" rel="stylesheet">
    <script src="https://cdn.quilljs.com/1.3.6/quill.js"></script>

    <div id="toolbar">

        <button class="ql-bold"></button>
        <button class="ql-italic"></button>
        <button class="ql-underline"></button>

        <button class="ql-list" value="ordered"></button>
        <button class="ql-list" value="bullet"></button>

        <button class="ql-link"></button>

        <button class="ql-clean"></button>

    </div>

    <div id="editor" style="
        background:white;
        color:black;
        height:320px;
        font-size:14px;
        font-family:Arial;
        padding:10px;
    ">
        <p>Hi,</p>

        <p>
            We're launching Dotpe Horizon, an AI-powered business intelligence digest,
            built entirely from your Rista data to help grow your revenue.
        </p>

        <p>
            Every week, on WhatsApp, you'll get:
        </p>

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

    </div>

    <script>

        var quill = new Quill('#editor', {

            theme: 'snow',

            modules: {
                toolbar: '#toolbar'
            }

        });

        document.querySelector('.ql-editor').style.minHeight = "300px";
        document.querySelector('.ql-editor').style.color = "black";
        document.querySelector('.ql-editor').style.backgroundColor = "white";
        document.querySelector('.ql-editor').style.caretColor = "black";

    </script>

    """,
    height=420,
    scrolling=True
)

# ---------------- ATTACHMENTS ---------------- #

attachments = st.file_uploader(
    "Attachments (Optional)",
    type=[
        "pdf",
        "doc",
        "docx",
        "xls",
        "xlsx",
        "csv",
        "jpg",
        "jpeg",
        "png"
    ],
    accept_multiple_files=True
)

# ---------------- SIGNATURE ---------------- #

signature = st.text_area(
    "Paste Email Signature Here",
    height=150
)

# ---------------- LOGO ---------------- #

logo_path = "dotpe_logo.png"

st.image(logo_path, width=180)

# ---------------- FUNCTIONS ---------------- #

def send_bulk_emails(
    sender_email,
    app_password,
    df,
    cc_emails,
    bcc_emails,
    subject,
    signature,
    attachments
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

            personalized_body = f"""
            <p>
                Hi {receiver_name.split()[0]},
            </p>

            <p>
                We're launching Dotpe Horizon, an AI-powered business intelligence digest,
                built entirely from your Rista data to help grow your revenue.
            </p>

            <p>
                Every week, on WhatsApp, you'll get:
            </p>

            <ul>
                <li>
                    What's actually driving (or dragging) your revenue
                </li>

                <li>
                    Where your orders, AOV, customers and margins moved - and why
                </li>

                <li>
                    Actions to address before next week
                </li>
            </ul>

            <p>
                It's private. It's yours. No benchmarks, no comparisons - just your numbers.
                To start receiving it, simply reply to this email with "YES".
                Your data is never shared with anyone outside Horizon.
            </p>

            <br>

            <p>
                Team Dotpe Horizon
            </p>
            """

            final_body = f"""
            <html>

            <body style="
                background-color:white;
                color:black;
                font-family:Arial;
                padding:20px;
            ">

                <div style="
                    font-size:14px;
                    line-height:1.6;
                ">

                    {personalized_body}

                    <br><br>

                    <b>
                        {signature.replace(chr(10), '<br>')}
                    </b>

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

            # ---------------- ATTACH LOGO ---------------- #

            with open(logo_path, "rb") as img:

                mime_img = MIMEImage(img.read())

                mime_img.add_header(
                    "Content-ID",
                    "<dotpelogo>"
                )

                msg.attach(mime_img)

            # ---------------- ATTACH FILES ---------------- #

            if attachments:

                for attachment in attachments:

                    mime_type, _ = mimetypes.guess_type(
                        attachment.name
                    )

                    if mime_type is None:
                        mime_type = "application/octet-stream"

                    main_type, sub_type = mime_type.split("/", 1)

                    part = MIMEBase(main_type, sub_type)

                    part.set_payload(
                        attachment.read()
                    )

                    encoders.encode_base64(part)

                    part.add_header(
                        "Content-Disposition",
                        f'attachment; filename="{attachment.name}"'
                    )

                    msg.attach(part)

            # ---------------- SEND EMAIL ---------------- #

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
                    signature,
                    attachments
                )

                st.success(
                    "Emails sent successfully!"
                )

        except Exception as e:

            st.error(f"Error: {str(e)}")
