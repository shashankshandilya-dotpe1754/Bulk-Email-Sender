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
