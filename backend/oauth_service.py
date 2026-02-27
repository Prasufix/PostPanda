from __future__ import annotations

import os
import time
from urllib.parse import urlencode

import requests


class OAuthError(Exception):
    pass


def provider_catalog() -> dict[str, dict[str, str | bool]]:
    microsoft_ready = bool(os.getenv("MS_CLIENT_ID") and os.getenv("MS_CLIENT_SECRET"))
    google_ready = bool(os.getenv("GOOGLE_CLIENT_ID") and os.getenv("GOOGLE_CLIENT_SECRET"))

    return {
        "microsoft": {
            "label": "Microsoft",
            "configured": microsoft_ready,
        },
        "google": {
            "label": "Google",
            "configured": google_ready,
        },
    }


def validate_provider(provider: str) -> str:
    normalized = (provider or "").strip().lower()
    if normalized not in ("microsoft", "google"):
        raise OAuthError("Unbekannter OAuth-Provider.")

    catalog = provider_catalog()
    if not catalog[normalized]["configured"]:
        raise OAuthError(
            f"OAuth fuer {catalog[normalized]['label']} ist nicht konfiguriert. Bitte CLIENT_ID und CLIENT_SECRET setzen."
        )

    return normalized


def build_authorization_url(provider: str, state: str, redirect_uri: str) -> str:
    provider = validate_provider(provider)

    if provider == "microsoft":
        tenant = os.getenv("MS_TENANT_ID", "common").strip() or "common"
        auth_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
        params = {
            "client_id": os.getenv("MS_CLIENT_ID", ""),
            "response_type": "code",
            "redirect_uri": redirect_uri,
            "response_mode": "query",
            "scope": "offline_access User.Read Mail.Send",
            "state": state,
        }
        return f"{auth_url}?{urlencode(params)}"

    auth_url = "https://accounts.google.com/o/oauth2/v2/auth"
    params = {
        "client_id": os.getenv("GOOGLE_CLIENT_ID", ""),
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "scope": "openid email https://www.googleapis.com/auth/gmail.send",
        "access_type": "offline",
        "prompt": "consent",
        "include_granted_scopes": "true",
        "state": state,
    }
    return f"{auth_url}?{urlencode(params)}"


def exchange_authorization_code(provider: str, code: str, redirect_uri: str) -> dict[str, str | float | None]:
    provider = validate_provider(provider)

    if provider == "microsoft":
        tenant = os.getenv("MS_TENANT_ID", "common").strip() or "common"
        token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
        payload = {
            "client_id": os.getenv("MS_CLIENT_ID", ""),
            "client_secret": os.getenv("MS_CLIENT_SECRET", ""),
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
            "scope": "offline_access User.Read Mail.Send",
        }
    else:
        token_url = "https://oauth2.googleapis.com/token"
        payload = {
            "client_id": os.getenv("GOOGLE_CLIENT_ID", ""),
            "client_secret": os.getenv("GOOGLE_CLIENT_SECRET", ""),
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
        }

    data = _post_form(token_url, payload)
    return _normalize_token_response(data)


def refresh_access_token(provider: str, refresh_token: str) -> dict[str, str | float | None]:
    provider = validate_provider(provider)

    if not refresh_token:
        raise OAuthError("OAuth-Token ist abgelaufen. Bitte erneut einloggen.")

    if provider == "microsoft":
        tenant = os.getenv("MS_TENANT_ID", "common").strip() or "common"
        token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
        payload = {
            "client_id": os.getenv("MS_CLIENT_ID", ""),
            "client_secret": os.getenv("MS_CLIENT_SECRET", ""),
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": "offline_access User.Read Mail.Send",
        }
    else:
        token_url = "https://oauth2.googleapis.com/token"
        payload = {
            "client_id": os.getenv("GOOGLE_CLIENT_ID", ""),
            "client_secret": os.getenv("GOOGLE_CLIENT_SECRET", ""),
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }

    data = _post_form(token_url, payload)
    normalized = _normalize_token_response(data)
    if not normalized.get("refresh_token"):
        normalized["refresh_token"] = refresh_token
    return normalized


def fetch_user_email(provider: str, access_token: str) -> str:
    provider = validate_provider(provider)

    if provider == "microsoft":
        response = requests.get(
            "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=20,
        )
        data = _read_json_response(response)
        email = (data.get("mail") or data.get("userPrincipalName") or "").strip()
    else:
        response = requests.get(
            "https://www.googleapis.com/oauth2/v2/userinfo",
            headers={"Authorization": f"Bearer {access_token}"},
            params={"alt": "json"},
            timeout=20,
        )
        data = _read_json_response(response)
        email = str(data.get("email", "")).strip()

    if not email:
        raise OAuthError("E-Mail-Adresse aus OAuth-Profil konnte nicht gelesen werden.")

    return email


def _normalize_token_response(data: dict[str, object]) -> dict[str, str | float | None]:
    access_token = str(data.get("access_token", "")).strip()
    refresh_token = data.get("refresh_token")
    refresh_token_str = str(refresh_token).strip() if isinstance(refresh_token, str) else None

    try:
        expires_in = int(data.get("expires_in", 3600))
    except (TypeError, ValueError):
        expires_in = 3600

    if not access_token:
        raise OAuthError("OAuth access_token fehlt in der Provider-Antwort.")

    return {
        "access_token": access_token,
        "refresh_token": refresh_token_str,
        "expires_at": time.time() + max(expires_in - 60, 60),
    }


def _post_form(url: str, payload: dict[str, str]) -> dict[str, object]:
    response = requests.post(url, data=payload, timeout=20)
    return _read_json_response(response)


def _read_json_response(response: requests.Response) -> dict[str, object]:
    data: dict[str, object]
    try:
        data = response.json()
    except Exception as exc:
        raise OAuthError(f"OAuth-Provider antwortet nicht mit JSON (HTTP {response.status_code}).") from exc

    if response.status_code >= 400:
        description = (
            str(data.get("error_description", "")).strip()
            or str(data.get("error", "")).strip()
            or f"HTTP {response.status_code}"
        )
        raise OAuthError(f"OAuth-Fehler: {description}")

    return data
