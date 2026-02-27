from __future__ import annotations

import json
import os
import time
from typing import Any

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, jsonify, make_response, redirect, request
from werkzeug.exceptions import HTTPException

from backend.mailer import (
    MailAppDraftConfig,
    MailConfig,
    MailerError,
    OAuthMailConfig,
    create_all_drafts_mail_app,
    create_test_draft_mail_app,
    is_valid_email,
    render_html_from_markup,
    render_template,
    send_all_mails,
    send_all_mails_oauth,
    send_test_mail,
    send_test_mail_oauth,
    to_text,
)
from backend.oauth_service import (
    OAuthError,
    build_authorization_url,
    exchange_authorization_code,
    fetch_user_email,
    provider_catalog,
    refresh_access_token,
    validate_provider,
)
from backend.oauth_store import OAuthAccount, oauth_store
from backend.store import store

load_dotenv()


class ApiError(Exception):
    def __init__(self, message: str, status_code: int = 400):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


def create_app() -> Flask:
    app = Flask(__name__)

    @app.get("/")
    def root() -> Any:
        return jsonify(
            {
                "ok": True,
                "message": "PostPanda backend is running.",
                "frontend": _default_frontend_origin(),
                "health": "/api/health",
                "oauthProviders": provider_catalog(),
            }
        )

    @app.errorhandler(ApiError)
    def handle_api_error(error: ApiError):
        return jsonify({"error": error.message}), error.status_code

    @app.errorhandler(MailerError)
    def handle_mailer_error(error: MailerError):
        return jsonify({"error": str(error)}), 400

    @app.errorhandler(OAuthError)
    def handle_oauth_error(error: OAuthError):
        return jsonify({"error": str(error)}), 400

    @app.errorhandler(HTTPException)
    def handle_http_exception(error: HTTPException):
        return (
            jsonify(
                {
                    "error": error.description,
                    "status": error.code,
                    "path": request.path,
                }
            ),
            error.code or 500,
        )

    @app.errorhandler(Exception)
    def handle_unexpected_error(error: Exception):
        return jsonify({"error": str(error)}), 500

    @app.get("/api/health")
    def health() -> Any:
        return jsonify({"ok": True})

    @app.get("/api/oauth/status")
    def oauth_status() -> Any:
        client_id = str(request.args.get("clientId", "")).strip()
        catalog = provider_catalog()
        accounts = oauth_store.list_accounts(client_id) if client_id else {}

        providers: dict[str, dict[str, Any]] = {}
        for provider, meta in catalog.items():
            account = accounts.get(provider)
            providers[provider] = {
                "label": meta["label"],
                "configured": bool(meta["configured"]),
                "connected": bool(account),
                "email": account.email if account else "",
                "expiresAt": account.expires_at if account else None,
            }

        return jsonify({"providers": providers})

    @app.get("/api/oauth/login/<provider>")
    def oauth_login(provider: str) -> Any:
        normalized_provider = validate_provider(provider)
        client_id = str(request.args.get("clientId", "")).strip()
        frontend_origin = str(request.args.get("frontendOrigin", "")).strip() or _default_frontend_origin()

        if not client_id:
            raise ApiError("clientId is required for OAuth.")

        state = oauth_store.create_pending_state(normalized_provider, client_id, frontend_origin)
        redirect_uri = _oauth_callback_uri(normalized_provider)
        authorization_url = build_authorization_url(normalized_provider, state, redirect_uri)

        return redirect(authorization_url, code=302)

    @app.get("/api/oauth/callback/<provider>")
    def oauth_callback(provider: str) -> Any:
        normalized_provider = _normalize_oauth_provider(provider)
        state = str(request.args.get("state", "")).strip()
        pending = oauth_store.consume_pending_state(state) if state else None

        frontend_origin = pending.frontend_origin if pending else _default_frontend_origin()

        if pending is None:
            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "error",
                    "provider": normalized_provider,
                    "message": "OAuth state is invalid or expired.",
                },
                status_code=400,
            )

        if pending.provider != normalized_provider:
            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "error",
                    "provider": normalized_provider,
                    "message": "OAuth provider does not match the login state.",
                },
                status_code=400,
            )

        if "error" in request.args:
            description = str(request.args.get("error_description", "")).strip() or str(request.args["error"])
            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "error",
                    "provider": normalized_provider,
                    "message": f"OAuth login canceled: {description}",
                },
                status_code=400,
            )

        code = str(request.args.get("code", "")).strip()
        if not code:
            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "error",
                    "provider": normalized_provider,
                    "message": "OAuth code is missing in the callback URL.",
                },
                status_code=400,
            )

        try:
            redirect_uri = _oauth_callback_uri(normalized_provider)
            token_data = exchange_authorization_code(normalized_provider, code, redirect_uri)
            access_token = str(token_data["access_token"])
            refresh_token = token_data.get("refresh_token")
            email = fetch_user_email(normalized_provider, access_token)

            account = OAuthAccount(
                provider=normalized_provider,
                email=email,
                access_token=access_token,
                refresh_token=str(refresh_token).strip() if isinstance(refresh_token, str) else None,
                expires_at=float(token_data["expires_at"]),
                updated_at=time.time(),
            )
            oauth_store.set_account(pending.client_id, account)

            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "ok",
                    "provider": normalized_provider,
                    "email": email,
                    "message": "OAuth connection successful.",
                },
            )
        except Exception as exc:
            return _oauth_popup_response(
                frontend_origin,
                {
                    "status": "error",
                    "provider": normalized_provider,
                    "message": str(exc),
                },
                status_code=400,
            )

    @app.post("/api/oauth/logout")
    def oauth_logout() -> Any:
        payload = _require_json(request.get_json(silent=True))
        client_id = str(payload.get("clientId", "")).strip()
        provider = _normalize_oauth_provider(str(payload.get("provider", "")))

        if not client_id:
            raise ApiError("clientId is required for OAuth.")

        oauth_store.remove_account(client_id, provider)
        return jsonify({"ok": True})

    @app.post("/api/upload")
    def upload_excel() -> Any:
        if "file" not in request.files:
            raise ApiError("No file was uploaded.")

        file = request.files["file"]
        if not file or not file.filename:
            raise ApiError("Please select an Excel file.")

        try:
            dataframe = pd.read_excel(file)
        except Exception as exc:
            raise ApiError(f"Failed to read Excel file: {exc}") from exc

        dataframe.columns = [str(col).strip() for col in dataframe.columns]
        dataframe = dataframe.dropna(how="all")

        if dataframe.empty:
            raise ApiError("The Excel file contains no recipient rows.")

        session_id = store.create(dataframe)

        return jsonify(
            {
                "sessionId": session_id,
                "filename": file.filename,
                "columns": list(dataframe.columns),
                "totalRows": len(dataframe),
            }
        )

    @app.post("/api/preview")
    def preview_message() -> Any:
        payload = _require_json(request.get_json(silent=True))
        dataframe = _require_dataframe(payload)
        mapping = _extract_mapping(payload)

        index = _normalize_index(payload.get("index", 0), len(dataframe))
        template = str(payload.get("template", ""))

        row = dataframe.iloc[index].to_dict()
        preview = render_template(
            row=row,
            template=template,
            email_col=mapping["email_col"],
            variable_map=mapping["variable_map"],
        )
        preview_html = render_html_from_markup(preview)

        recipient = to_text(row.get(mapping["email_col"], "")) if mapping["email_col"] else ""

        return jsonify(
            {
                "index": index,
                "totalRows": len(dataframe),
                "recipient": recipient,
                "recipientValid": is_valid_email(recipient),
                "preview": preview,
                "previewHtml": preview_html,
            }
        )

    @app.post("/api/send-test")
    def send_test() -> Any:
        payload = _require_json(request.get_json(silent=True))
        dataframe = _require_dataframe(payload)
        auth_mode = _extract_auth_mode(payload)

        index = _normalize_index(payload.get("index", 0), len(dataframe))

        if auth_mode == "oauth":
            config = _extract_oauth_mail_config(payload)
            send_test_mail_oauth(dataframe, index, config)
            message = f"Test email to {config.sender} was sent."
        elif auth_mode == "mailapp":
            config = _extract_mail_app_config(payload)
            recipient = create_test_draft_mail_app(dataframe, index, config)
            message = f"Message for {recipient} was opened in the default mail app."
        else:
            config = _extract_mail_config(payload)
            send_test_mail(dataframe, index, config)
            message = f"Test email to {config.sender} was sent."

        return jsonify({"ok": True, "mode": auth_mode, "message": message})

    @app.post("/api/send-all")
    def send_all() -> Any:
        payload = _require_json(request.get_json(silent=True))
        dataframe = _require_dataframe(payload)
        auth_mode = _extract_auth_mode(payload)

        if auth_mode == "oauth":
            config = _extract_oauth_mail_config(payload)
            result = send_all_mails_oauth(dataframe, config)
            return jsonify({"ok": True, "mode": "oauth", **result})

        if auth_mode == "mailapp":
            config = _extract_mail_app_config(payload)
            result = create_all_drafts_mail_app(dataframe, config)
            return jsonify(
                {
                    "ok": True,
                    "mode": "mailapp",
                    "total": result["total"],
                    "sent": result["drafted"],
                    "drafted": result["drafted"],
                    "skipped": result["skipped"],
                }
            )

        config = _extract_mail_config(payload)
        result = send_all_mails(dataframe, config)
        return jsonify({"ok": True, "mode": "password", **result})

    return app


def _require_json(payload: dict[str, Any] | None) -> dict[str, Any]:
    if payload is None:
        raise ApiError("Invalid JSON body.")
    return payload


def _require_dataframe(payload: dict[str, Any]) -> pd.DataFrame:
    session_id = str(payload.get("sessionId", "")).strip()
    if not session_id:
        raise ApiError("sessionId is required.")

    dataframe = store.get(session_id)
    if dataframe is None:
        raise ApiError("Session not found. Please upload the Excel file again.", 404)
    if dataframe.empty:
        raise ApiError("The recipient list is empty.")
    return dataframe


def _normalize_index(value: Any, total: int) -> int:
    if total <= 0:
        return 0
    try:
        index = int(value)
    except (TypeError, ValueError):
        index = 0
    return index % total


def _extract_auth_mode(payload: dict[str, Any]) -> str:
    mode = str(payload.get("authMode", "password")).strip().lower()
    if mode not in ("password", "oauth", "mailapp"):
        raise ApiError("authMode must be 'password', 'oauth', or 'mailapp'.")
    return mode


def _extract_mapping(payload: dict[str, Any]) -> dict[str, Any]:
    mapping = payload.get("mapping") or {}
    raw_variable_map = mapping.get("variableMap") or {}

    if not isinstance(raw_variable_map, dict):
        raise ApiError("mapping.variableMap must be an object.")

    variable_map: dict[str, str] = {}
    for variable_name, column_name in raw_variable_map.items():
        name = str(variable_name).strip().strip("{}").strip()
        column = str(column_name).strip()
        if not name or not column:
            continue
        if name.lower() == "mail":
            raise ApiError("Variable 'Mail' is reserved. Please use a different name.")
        variable_map[name] = column

    legacy_firstname_col = str(mapping.get("firstnameCol", "")).strip()
    legacy_name_col = str(mapping.get("nameCol", "")).strip()
    if legacy_firstname_col:
        variable_map.setdefault("FirstName", legacy_firstname_col)
        variable_map.setdefault("Vorname", legacy_firstname_col)
    if legacy_name_col:
        variable_map.setdefault("LastName", legacy_name_col)
        variable_map.setdefault("Name", legacy_name_col)

    return {
        "email_col": str(mapping.get("emailCol", "")).strip(),
        "variable_map": variable_map,
    }


def _extract_message_fields(payload: dict[str, Any]) -> tuple[dict[str, Any], str, str]:
    mapping = _extract_mapping(payload)
    subject = str(payload.get("subject", "")).strip()
    template = str(payload.get("template", ""))

    if not subject:
        raise ApiError("Please enter a subject.")
    if not template.strip():
        raise ApiError("Please enter a message body.")
    if not mapping["email_col"]:
        raise ApiError("Please select an email column.")

    return mapping, subject, template


def _extract_mail_config(payload: dict[str, Any]) -> MailConfig:
    mapping, subject, template = _extract_message_fields(payload)
    smtp = payload.get("smtp") or {}

    sender = str(smtp.get("sender", "")).strip()
    password = str(smtp.get("password", "")).strip()
    host = str(smtp.get("host", "")).strip()

    try:
        port = int(smtp.get("port", 0))
    except (TypeError, ValueError):
        raise ApiError("SMTP port is invalid.")

    if not is_valid_email(sender):
        raise ApiError("Please enter a valid sender email.")
    if not password:
        raise ApiError("Please enter an SMTP password.")
    if not host:
        raise ApiError("Please enter an SMTP host.")
    if port < 1 or port > 65535:
        raise ApiError("SMTP port must be between 1 and 65535.")

    return MailConfig(
        sender=sender,
        password=password,
        host=host,
        port=port,
        subject=subject,
        template=template,
        email_col=mapping["email_col"],
        variable_map=mapping["variable_map"],
    )


def _extract_mail_app_config(payload: dict[str, Any]) -> MailAppDraftConfig:
    mapping, subject, template = _extract_message_fields(payload)
    mail_app = payload.get("mailApp") or {}
    provider = str(mail_app.get("provider", "")).strip().lower()
    if provider not in ("outlook", "gmail", "custom"):
        provider = "custom"
    return MailAppDraftConfig(
        provider=provider,
        subject=subject,
        template=template,
        email_col=mapping["email_col"],
        variable_map=mapping["variable_map"],
    )


def _extract_oauth_mail_config(payload: dict[str, Any]) -> OAuthMailConfig:
    mapping, subject, template = _extract_message_fields(payload)
    oauth = payload.get("oauth") or {}

    provider = _normalize_oauth_provider(str(oauth.get("provider", "")))
    client_id = str(oauth.get("clientId", "")).strip()

    if not client_id:
        raise ApiError("clientId is required for OAuth.")

    account = oauth_store.get_account(client_id, provider)
    if account is None:
        raise ApiError("No OAuth login found. Please connect first.")

    account = _refresh_oauth_account_if_needed(account)
    oauth_store.set_account(client_id, account)

    return OAuthMailConfig(
        provider=provider,
        sender=account.email,
        access_token=account.access_token,
        subject=subject,
        template=template,
        email_col=mapping["email_col"],
        variable_map=mapping["variable_map"],
    )


def _refresh_oauth_account_if_needed(account: OAuthAccount) -> OAuthAccount:
    if not account.is_expired():
        return account

    token_data = refresh_access_token(account.provider, account.refresh_token or "")
    refreshed_token = token_data.get("refresh_token")

    return OAuthAccount(
        provider=account.provider,
        email=account.email,
        access_token=str(token_data["access_token"]),
        refresh_token=str(refreshed_token).strip() if isinstance(refreshed_token, str) else account.refresh_token,
        expires_at=float(token_data["expires_at"]),
        updated_at=time.time(),
    )


def _normalize_oauth_provider(provider: str) -> str:
    normalized = provider.strip().lower()
    if normalized not in ("microsoft", "google"):
        raise ApiError("OAuth provider must be 'microsoft' or 'google'.")
    return normalized


def _oauth_callback_base() -> str:
    base = os.getenv("POSTPANDA_OAUTH_CALLBACK_BASE")
    return (base or "http://127.0.0.1:8000").rstrip("/")


def _oauth_callback_uri(provider: str) -> str:
    return f"{_oauth_callback_base()}/api/oauth/callback/{provider}"


def _default_frontend_origin() -> str:
    origin = os.getenv("POSTPANDA_FRONTEND_ORIGIN")
    return (origin or "http://127.0.0.1:5173").rstrip("/")


def _oauth_popup_response(frontend_origin: str, payload: dict[str, Any], status_code: int = 200):
    safe_origin = frontend_origin if frontend_origin.startswith(("http://", "https://")) else _default_frontend_origin()
    post_message_payload = {"type": "postpanda-oauth", **payload}

    html = f"""<!doctype html>
<html>
  <head>
    <meta charset=\"utf-8\" />
    <title>PostPanda OAuth</title>
    <style>
      body {{ font-family: Arial, sans-serif; padding: 20px; color: #1f2937; }}
      .hint {{ color: #6b7280; font-size: 14px; }}
    </style>
  </head>
  <body>
    <h3>OAuth completed</h3>
    <p class=\"hint\">This window can now be closed.</p>
    <script>
      const payload = {json.dumps(post_message_payload)};
      const targetOrigin = {json.dumps(safe_origin)};
      if (window.opener) {{
        window.opener.postMessage(payload, targetOrigin);
      }}
      setTimeout(() => window.close(), 120);
    </script>
  </body>
</html>
"""
    response = make_response(html, status_code)
    response.headers["Content-Type"] = "text/html; charset=utf-8"
    return response


app = create_app()


if __name__ == "__main__":
    port_value = os.getenv("POSTPANDA_BACKEND_PORT", "8000")
    port = int(port_value)
    app.run(host="127.0.0.1", port=port, debug=True)
