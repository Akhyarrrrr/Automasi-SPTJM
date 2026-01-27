from __future__ import annotations

import os
import smtplib
import time
from email.message import EmailMessage


def send_smtp(
    *,
    to_email: str,
    subject: str,
    body: str,
    attachment_bytes: bytes | None = None,
    attachment_filename: str | None = None,
    sleep_seconds: float = 0.7,
    timeout: int = 30,
    retries: int = 2,
):
    """
    Kirim email via SMTP server (generic, bukan Gmail-hardcode).
    Ambil konfigurasi dari env:
      SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM_NAME (opsional)
    """

    SMTP_HOST = os.getenv("SMTP_HOST")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")
    FROM_NAME = os.getenv("SMTP_FROM_NAME", SMTP_USER or "")

    if not all([SMTP_HOST, SMTP_USER, SMTP_PASS]):
        raise RuntimeError(
            "SMTP belum dikonfigurasi lengkap di .env (SMTP_HOST/SMTP_USER/SMTP_PASS)."
        )

    msg = EmailMessage()
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>".strip()
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_bytes and attachment_filename:
        msg.add_attachment(
            attachment_bytes,
            maintype="application",
            subtype="pdf",
            filename=attachment_filename,
        )

    last_err = None
    for attempt in range(retries + 1):
        try:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=timeout) as smtp:
                # STARTTLS umum untuk port 587
                smtp.starttls()
                smtp.login(SMTP_USER, SMTP_PASS)
                smtp.send_message(msg)

            time.sleep(sleep_seconds)
            return

        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(1.2 * (attempt + 1))
            else:
                raise last_err
