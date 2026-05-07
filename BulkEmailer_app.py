# streamlit_app.py

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

# ---------------- CLEAR QUILL EMAIL BODY ---------------- #

# ---------------- EMAIL BODY ---------------- #

email_body = st.text_area(
    "Email Body",
    height=320
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

        # ---------------- NORMAL RECEIVER EMAILS ---------------- #

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

        # ---------------- EXCEL BCC RECEIVERS ---------------- #

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

        # ---------------- FINAL RECEIVER LIST ---------------- #

        final_receivers = manual_receivers + excel_receivers

        for person in final_receivers:

            receiver_email = person["email"]
            receiver_name = person["name"]

            personalized_body = f"""
            <p>
                Hi {receiver_name},
            </p>
            
            {email_body}
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

                bcc_list = []
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

# ---------------- LOAD LOCAL IMAGES ---------------- #

def get_base64(image_path):

    with open(image_path, "rb") as img_file:

        return base64.b64encode(
            img_file.read()
        ).decode()

background_image = get_base64(
    "maxbulk-using-gmail.png"
)

logo_image = get_base64(
    "dotpe_logo.png"
)

# ---------------- CUSTOM CSS ---------------- #

st.markdown(f"""
<style>

    .stApp {{
        background: linear-gradient(
            rgba(0, 0, 0, 0.72),
            rgba(0, 0, 0, 0.72)
        ),
        url("data:image/png;base64,{background_image}");

        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        color: white;
    }}

    .main > div {{
        background: rgba(17, 17, 17, 0.78);
        padding: 30px;
        border-radius: 22px;
        backdrop-filter: blur(10px);
        box-shadow: 0px 0px 35px rgba(0,0,0,0.55);
        margin-top: 25px;
    }}

    .top-logo {{
        position: fixed;
        top: 12px;
        left: 20px;
        z-index: 9999;
    }}

    .top-logo img {{
        width: 170px;
        border-radius: 10px;
        background: white;
        padding: 6px 10px;
        box-shadow: 0px 4px 18px rgba(0,0,0,0.45);
    }}

    .main-title {{
        text-align: center;
        color: white;
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 30px;
        letter-spacing: 0.5px;
        text-shadow: 2px 2px 15px rgba(0,0,0,0.7);
    }}

    .block-container {{
        padding-top: 2rem;
    }}

    .stTextInput input {{
        background: rgba(30, 30, 30, 0.78) !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 12px !important;
        height: 52px;
        transition: 0.3s ease;
    }}

    .stTextInput input:focus {{
        border: 1px solid #ff4b4b !important;
        box-shadow: 0px 0px 12px rgba(255,75,75,0.55);
    }}

    .stTextArea textarea {{
        background: rgba(30, 30, 30, 0.78) !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 12px !important;
        transition: 0.3s ease;
    }}

    .stTextArea textarea:focus {{
        border: 1px solid #ff4b4b !important;
        box-shadow: 0px 0px 12px rgba(255,75,75,0.55);
    }}

    input::placeholder,
    textarea::placeholder {{
        color: #d9d9d9 !important;
        opacity: 1 !important;
    }}

    .stFileUploader {{
        background: rgba(30, 30, 30, 0.78) !important;
        padding: 14px;
        border-radius: 16px;
        border: 1px solid rgba(255,255,255,0.12);
        color: white !important;
        transition: 0.3s ease;
    }}

    section[data-testid="stFileUploaderDropzone"] {{
        background: rgba(30, 30, 30, 0.78) !important;
        color: white !important;
        border: 2px dashed rgba(255,255,255,0.2) !important;
        border-radius: 16px !important;
        transition: 0.3s ease;
    }}

    section[data-testid="stFileUploaderDropzone"]:hover {{
        border: 2px dashed #ff4b4b !important;
        background: rgba(50,50,50,0.88) !important;
    }}

    section[data-testid="stFileUploaderDropzone"] * {{
        color: white !important;
    }}

    label {{
        color: white !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        letter-spacing: 0.3px;
    }}

    .stButton button {{
        width: 100%;
        background: linear-gradient(
            135deg,
            #ff4b4b,
            #d32f2f
        );

        color: white;
        font-weight: 700;
        font-size: 17px;
        border-radius: 14px;
        height: 56px;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0px 6px 20px rgba(255,75,75,0.35);
    }}

    .stButton button:hover {{
        transform: translateY(-2px);
        background: linear-gradient(
            135deg,
            #ff5c5c,
            #b71c1c
        );

        box-shadow: 0px 8px 24px rgba(255,75,75,0.5);
        color: white;
    }}

    .stAlert {{
        border-radius: 14px !important;
    }}

</style>
""", unsafe_allow_html=True)

# ---------------- TOP LEFT LOGO ---------------- #

st.markdown(
    f"""
    <div class="top-logo">
        <img src="data:image/png;base64,{logo_image}">
    </div>
    """,
    unsafe_allow_html=True
)

# ---------------- FIX TOP WHITE BAR ---------------- #

st.markdown("""
<style>

    .block-container {
        padding-top: 1rem !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0px !important;
    }

    .stApp > header {
        background-color: transparent !important;
    }

    .main {
        padding-top: 0rem !important;
    }

    .top-logo {
        position: fixed;
        top: 8px;
        left: 18px;
        z-index: 999999;
    }

    .top-logo img {
        width: 170px;
        border-radius: 10px;
        background: white;
        padding: 6px 10px;
        box-shadow: 0px 4px 18px rgba(0,0,0,0.45);
    }

</style>
""", unsafe_allow_html=True)

# ---------------- CLEAR ALL FIELDS ON REFRESH ---------------- #

if "clear_state" not in st.session_state:

    st.session_state.clear_state = True

    for key in list(st.session_state.keys()):

        del st.session_state[key]

st.set_page_config(
    page_title="Bulk Email Sender",
    layout="centered"
)
