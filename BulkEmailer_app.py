import streamlit as st
import pandas as pd
import smtplib
import ssl
import re
import mimetypes
import base64

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from streamlit_quill import st_quill

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

receiver_emails = st.text_area(
    "Receiver Emails: (Optional - emails separated by comma)",
    placeholder="xyz@email.com, abc@email.com"
)

cc_emails = st.text_area(
    "Cc: (emails separated by comma)",
    placeholder="xyz@email.com, abc@email.com"
)

uploaded_file = st.file_uploader(
    "Bcc Excel File (Upload Excel File)",
    type=["xlsx", "xls"]
)

subject = st.text_input("Subject")

# ---------------- EMAIL BODY ---------------- #

email_body = st_quill(
    value="",
    html=True,
    placeholder="Write your email body here...",
    key="email_editor"
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
    receiver_emails,
    bcc_list,
    cc_emails,
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

        manual_receivers = []

        if receiver_emails.strip():

            manual_receivers = [
                {
                    "email": email.strip(),
                    "name": email.strip().split("@")[0]
                }
                for email in receiver_emails.split(",")
                if email.strip()
            ]

        excel_receivers = []

        if bcc_list:

            for item in bcc_list:

                if isinstance(item, dict):

                    excel_receivers.append(item)

                else:

                    excel_receivers.append({
                        "email": str(item).strip(),
                        "name": str(item).split("@")[0]
                    })

        final_receivers = manual_receivers + excel_receivers

        for person in final_receivers:

            receiver_email = person["email"]
            receiver_name = person["name"]

            # ---------------- PERSONALIZED GREETING ---------------- #

            email_content = email_body.strip() if email_body else ""

            greetings = [
                "hi",
                "hello",
                "dear",
                "good morning",
                "good afternoon",
                "good evening"
            ]

            clean_text = re.sub(
                r"<[^>]+>",
                "",
                email_content
            ).strip().lower()

            greeting_found = False

            for greeting in greetings:

                if clean_text.startswith(greeting):

                    greeting_found = True
                    break

            if greeting_found:

                email_content = re.sub(
                    r"^(<p>)?(Hi|Hello|Dear|Good Morning|Good Afternoon|Good Evening)(\s|&nbsp;)*",
                    f"\\1\\2 {receiver_name}, ",
                    email_content,
                    flags=re.IGNORECASE
                )

                personalized_body = email_content

            else:

                personalized_body = f"""
                <p style="margin:0;">
                    Hi {receiver_name},
                </p>

                {email_content}
                """

            final_body = f"""
            <html>

            <body style="
                background-color:white;
                color:black;
                font-family:Arial;
                padding:10px 20px;
                margin:0;
            ">

                <div style="
                    font-size:14px;
                    line-height:1.2;
                    margin:0;
                    padding:0;
                ">

                    {personalized_body}

                    <br>

                    <div style="margin-top:10px;">
                        <b>
                            {signature.replace(chr(10), '<br>')}
                        </b>
                    </div>

                    <br>

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

                    attachment.seek(0)

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
    cc_emails,
    subject,
    signature
)

# ---------------- SEND BUTTON ---------------- #

if st.button("Send Emails"):

    if not all(mandatory_fields):

        missing_fields = []

        if not sender_email:
            missing_fields.append("Sender Email")

        if not app_password:
            missing_fields.append("Email App Password")

        if not cc_emails:
            missing_fields.append("CC Emails")

        if not subject:
            missing_fields.append("Subject")

        if not signature:
            missing_fields.append("Signature")

        st.error(
            f"Please fill all mandatory fields. ({', '.join(missing_fields)})"
        )

    else:

        try:

            bcc_list = []

            if uploaded_file:

                df = pd.read_excel(uploaded_file)

                if "Email" not in df.columns:

                    st.error(
                        "Excel file must contain Email column."
                    )

                    st.stop()

                for _, row in df.iterrows():

                    bcc_list.append({
                        "email": str(row["Email"]).strip(),
                        "name": str(row["Name"]).strip()
                        if "Name" in df.columns
                        else str(row["Email"]).split("@")[0]
                    })

            send_bulk_emails(
                sender_email,
                app_password,
                receiver_emails,
                bcc_list,
                cc_emails,
                subject,
                signature,
                attachments
            )

            st.success(
                "Emails sent successfully!"
            )

        except Exception as e:

            st.error(f"Error: {str(e)}")
