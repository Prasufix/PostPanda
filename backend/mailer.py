from __future__ import annotations

import base64
import html
import re
import smtplib
import ssl
import subprocess
import sys
from urllib.parse import quote
from dataclasses import dataclass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any

import pandas as pd
import requests

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
BOLD_REGEX = re.compile(r"\*\*(.+?)\*\*")
LINK_REGEX = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


class MailerError(Exception):
    pass


@dataclass
class MailConfig:
    sender: str
    password: str
    host: str
    port: int
    subject: str
    template: str
    email_col: str
    variable_map: dict[str, str]


@dataclass
class OAuthMailConfig:
    provider: str
    sender: str
    access_token: str
    subject: str
    template: str
    email_col: str
    variable_map: dict[str, str]


@dataclass
class MailAppDraftConfig:
    provider: str
    subject: str
    template: str
    email_col: str
    variable_map: dict[str, str]


def to_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def is_valid_email(value: str) -> bool:
    return bool(EMAIL_REGEX.match((value or "").strip()))


def _safe_link(url: str) -> str | None:
    candidate = url.strip()
    if candidate.startswith(("http://", "https://", "mailto:")):
        return candidate
    return None


def render_html_from_markup(text: str) -> str:
    escaped = html.escape(text or "")

    def bold_replace(match: re.Match[str]) -> str:
        return f"<strong>{match.group(1)}</strong>"

    def link_replace(match: re.Match[str]) -> str:
        label = match.group(1).strip()
        raw_url = html.unescape(match.group(2).strip())
        safe_url = _safe_link(raw_url)
        if not safe_url:
            return match.group(0)
        href = html.escape(safe_url, quote=True)
        return f'<a href="{href}" target="_blank" rel="noopener noreferrer">{label}</a>'

    rendered = BOLD_REGEX.sub(bold_replace, escaped)
    rendered = LINK_REGEX.sub(link_replace, rendered)
    rendered = rendered.replace("\n", "<br>\n")
    return rendered


def render_plain_text_from_markup(text: str) -> str:
    plain = text or ""
    plain = LINK_REGEX.sub(lambda m: f"{m.group(1)} ({m.group(2)})", plain)
    plain = BOLD_REGEX.sub(lambda m: m.group(1), plain)
    return plain


def render_template(
    row: dict[str, Any],
    template: str,
    email_col: str,
    variable_map: dict[str, str] | None = None,
) -> str:
    body = template or ""
    email = to_text(row.get(email_col, "")) if email_col else ""
    replacements: dict[str, str] = {}
    normalized_replacements: dict[str, str] = {}

    def normalize_key(value: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", value.lower())

    def register_replacement(key: str, value: str) -> None:
        if not key:
            return
        replacements[key] = value
        normalized = normalize_key(key)
        if normalized and normalized not in normalized_replacements:
            normalized_replacements[normalized] = value

    # Direct Excel column placeholders: {{ColumnName}}
    for key, value in row.items():
        column_key = str(key).strip()
        if column_key:
            register_replacement(column_key, to_text(value))

    # Variable mapping placeholders: {{Variable}}
    if variable_map:
        for variable_name, column_name in variable_map.items():
            placeholder = str(variable_name).strip()
            column = str(column_name).strip()
            if not placeholder:
                continue
            register_replacement(placeholder, to_text(row.get(column, "")))

    # Fixed email placeholders
    register_replacement("Mail", email)
    register_replacement("Email", email)

    lower_replacements = {key.lower(): value for key, value in replacements.items()}

    def replace_match(match: re.Match[str]) -> str:
        token = match.group(1).strip()
        if token in replacements:
            return replacements[token]
        lower_token = token.lower()
        if lower_token in lower_replacements:
            return lower_replacements[lower_token]
        normalized_token = normalize_key(token)
        if normalized_token in normalized_replacements:
            return normalized_replacements[normalized_token]
        return match.group(0)

    return re.sub(r"{{\s*([^{}]+?)\s*}}", replace_match, body)


def build_message(config: MailConfig, row: dict[str, Any], recipient: str) -> MIMEMultipart:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = config.subject
    msg["From"] = config.sender
    msg["To"] = recipient

    body = render_template(
        row=row,
        template=config.template,
        email_col=config.email_col,
        variable_map=config.variable_map,
    )
    body_plain = render_plain_text_from_markup(body)
    body_html = render_html_from_markup(body)
    msg.attach(MIMEText(body_plain, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    return msg


def build_oauth_message(config: OAuthMailConfig, row: dict[str, Any], recipient: str) -> MIMEMultipart:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = config.subject
    msg["From"] = config.sender
    msg["To"] = recipient

    body = render_template(
        row=row,
        template=config.template,
        email_col=config.email_col,
        variable_map=config.variable_map,
    )
    body_plain = render_plain_text_from_markup(body)
    body_html = render_html_from_markup(body)
    msg.attach(MIMEText(body_plain, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    return msg


def connect_smtp(config: MailConfig) -> smtplib.SMTP:
    context = ssl.create_default_context()
    server = smtplib.SMTP(config.host, config.port, timeout=30)
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    server.login(config.sender, config.password)
    return server


def send_test_mail(dataframe: pd.DataFrame, index: int, config: MailConfig) -> None:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    row_dict = dataframe.iloc[index % len(dataframe)].to_dict()
    message = build_message(config, row_dict, config.sender)

    server = None
    try:
        server = connect_smtp(config)
        server.sendmail(config.sender, config.sender, message.as_string())
    finally:
        if server is not None:
            server.quit()


def send_all_mails(dataframe: pd.DataFrame, config: MailConfig) -> dict[str, int]:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    sent = 0
    skipped = 0
    total = len(dataframe)

    server = None
    try:
        server = connect_smtp(config)
        for _, row in dataframe.iterrows():
            row_dict = row.to_dict()
            recipient = to_text(row_dict.get(config.email_col, ""))
            if not is_valid_email(recipient):
                skipped += 1
                continue
            message = build_message(config, row_dict, recipient)
            server.sendmail(config.sender, recipient, message.as_string())
            sent += 1
    finally:
        if server is not None:
            server.quit()

    return {"total": total, "sent": sent, "skipped": skipped}


def send_test_mail_oauth(dataframe: pd.DataFrame, index: int, config: OAuthMailConfig) -> None:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    row_dict = dataframe.iloc[index % len(dataframe)].to_dict()
    _send_oauth_message(config, row_dict, config.sender)


def send_all_mails_oauth(dataframe: pd.DataFrame, config: OAuthMailConfig) -> dict[str, int]:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    sent = 0
    skipped = 0
    total = len(dataframe)

    for _, row in dataframe.iterrows():
        row_dict = row.to_dict()
        recipient = to_text(row_dict.get(config.email_col, ""))
        if not is_valid_email(recipient):
            skipped += 1
            continue
        _send_oauth_message(config, row_dict, recipient)
        sent += 1

    return {"total": total, "sent": sent, "skipped": skipped}


def create_test_draft_mail_app(dataframe: pd.DataFrame, index: int, config: MailAppDraftConfig) -> str:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    row_dict = dataframe.iloc[index % len(dataframe)].to_dict()
    recipient = to_text(row_dict.get(config.email_col, ""))
    if not is_valid_email(recipient):
        raise MailerError("Recipient for test draft is invalid.")

    body_markup = render_template(
        row=row_dict,
        template=config.template,
        email_col=config.email_col,
        variable_map=config.variable_map,
    )
    body_html = render_html_from_markup(body_markup)
    body_plain = render_plain_text_from_markup(body_markup)
    _create_mail_program_draft(config.provider, config.subject, recipient, body_plain, body_html)
    return recipient


def create_all_drafts_mail_app(dataframe: pd.DataFrame, config: MailAppDraftConfig) -> dict[str, int]:
    if dataframe.empty:
        raise MailerError("The recipient list is empty.")

    total = len(dataframe)
    drafted = 0
    skipped = 0

    for _, row in dataframe.iterrows():
        row_dict = row.to_dict()
        recipient = to_text(row_dict.get(config.email_col, ""))
        if not is_valid_email(recipient):
            skipped += 1
            continue

        body_markup = render_template(
            row=row_dict,
            template=config.template,
            email_col=config.email_col,
            variable_map=config.variable_map,
        )
        body_html = render_html_from_markup(body_markup)
        body_plain = render_plain_text_from_markup(body_markup)
        _create_mail_program_draft(config.provider, config.subject, recipient, body_plain, body_html)
        drafted += 1

    return {"total": total, "drafted": drafted, "skipped": skipped}


def _send_oauth_message(config: OAuthMailConfig, row: dict[str, Any], recipient: str) -> None:
    provider = config.provider.strip().lower()
    if provider == "microsoft":
        _send_with_microsoft_graph(config, row, recipient)
        return
    if provider == "google":
        _send_with_gmail_api(config, row, recipient)
        return
    raise MailerError("Unknown OAuth provider for sending.")


def _send_with_microsoft_graph(config: OAuthMailConfig, row: dict[str, Any], recipient: str) -> None:
    body_markup = render_template(
        row=row,
        template=config.template,
        email_col=config.email_col,
        variable_map=config.variable_map,
    )
    body_html = render_html_from_markup(body_markup)

    payload = {
        "message": {
            "subject": config.subject,
            "body": {
                "contentType": "HTML",
                "content": body_html,
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient,
                    }
                }
            ],
        },
        "saveToSentItems": "true",
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        json=payload,
        headers={
            "Authorization": f"Bearer {config.access_token}",
            "Content-Type": "application/json",
        },
        timeout=30,
    )

    if response.status_code >= 400:
        detail = _response_detail(response)
        raise MailerError(f"Microsoft Graph sendMail failed: {detail}")


def _send_with_gmail_api(config: OAuthMailConfig, row: dict[str, Any], recipient: str) -> None:
    message = build_oauth_message(config, row, recipient)
    raw_bytes = message.as_bytes()
    raw_base64 = base64.urlsafe_b64encode(raw_bytes).decode("ascii").rstrip("=")

    response = requests.post(
        "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
        json={"raw": raw_base64},
        headers={"Authorization": f"Bearer {config.access_token}"},
        timeout=30,
    )

    if response.status_code >= 400:
        detail = _response_detail(response)
        raise MailerError(f"Gmail API send failed: {detail}")


def _response_detail(response: requests.Response) -> str:
    try:
        data = response.json()
        if isinstance(data, dict):
            if "error" in data and isinstance(data["error"], dict):
                msg = str(data["error"].get("message", "")).strip()
                if msg:
                    return msg
            msg = str(data.get("error_description", "")).strip()
            if msg:
                return msg
            msg = str(data.get("error", "")).strip()
            if msg:
                return msg
    except Exception:
        pass
    text = response.text.strip()
    return text or f"HTTP {response.status_code}"


def _create_mail_program_draft(provider: str, subject: str, recipient: str, body_plain: str, body_html: str) -> None:
    normalized_provider = (provider or "").strip().lower()
    if normalized_provider == "outlook":
        _create_outlook_draft(subject, recipient, body_html)
        return
    _open_in_default_mail_app(subject, recipient, body_plain)


def _open_in_default_mail_app(subject: str, recipient: str, body: str) -> None:
    if sys.platform == "darwin":
        url = _build_mailto_url(recipient, subject, body)
        try:
            subprocess.run(["open", url], check=True, capture_output=True, text=True)
        except FileNotFoundError as exc:
            raise MailerError("'open' was not found on macOS.") from exc
        except subprocess.CalledProcessError as exc:
            detail = (exc.stderr or exc.stdout or "").strip()
            raise MailerError(f"Could not open mail app: {detail or 'Unknown error'}") from exc
        return

    if sys.platform.startswith("linux"):
        url = _build_mailto_url(recipient, subject, body)
        try:
            subprocess.run(["xdg-open", url], check=True, capture_output=True, text=True)
        except Exception as exc:
            raise MailerError(f"Could not open mail app: {exc}") from exc
        return

    raise MailerError("Mail app integration is not available on this operating system.")


def _build_mailto_url(recipient: str, subject: str, body: str) -> str:
    to_part = quote(recipient, safe="@._-+")
    subject_part = quote(subject or "", safe="")
    body_part = quote(body or "", safe="")
    return f"mailto:{to_part}?subject={subject_part}&body={body_part}"


def _create_outlook_draft(subject: str, recipient: str, html_content: str) -> None:
    if sys.platform != "darwin":
        raise MailerError("Outlook drafts via AppleScript are only available on macOS.")

    message_html = f"<html><body>{html_content}</body></html>"
    script = """
on run argv
  set msgSubject to item 1 of argv
  set msgRecipient to item 2 of argv
  set msgHtml to item 3 of argv

  tell application "Microsoft Outlook"
    activate
    set newMessage to make new outgoing message with properties {subject:msgSubject, content:msgHtml}
    make new recipient at newMessage with properties {email address:{address:msgRecipient}}
    open newMessage
  end tell
end run
"""

    try:
        subprocess.run(
            ["osascript", "-e", script, subject, recipient, message_html],
            check=True,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as exc:
        raise MailerError("osascript was not found.") from exc
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        if "Not authorized" in detail:
            detail = "Please allow Automation permission for Terminal/Python -> Microsoft Outlook in macOS."
        raise MailerError(f"Could not create Outlook draft: {detail or 'Unknown error'}") from exc
