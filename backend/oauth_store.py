from __future__ import annotations

import secrets
import time
from dataclasses import dataclass
from threading import Lock


@dataclass
class OAuthAccount:
    provider: str
    email: str
    access_token: str
    refresh_token: str | None
    expires_at: float
    updated_at: float

    def is_expired(self, skew_seconds: int = 60) -> bool:
        return self.expires_at <= (time.time() + skew_seconds)


@dataclass
class PendingOAuthState:
    provider: str
    client_id: str
    frontend_origin: str
    created_at: float


class OAuthStore:
    def __init__(self) -> None:
        self._accounts: dict[tuple[str, str], OAuthAccount] = {}
        self._pending_states: dict[str, PendingOAuthState] = {}
        self._lock = Lock()

    def set_account(self, client_id: str, account: OAuthAccount) -> None:
        key = (client_id, account.provider)
        with self._lock:
            self._accounts[key] = account

    def get_account(self, client_id: str, provider: str) -> OAuthAccount | None:
        key = (client_id, provider)
        with self._lock:
            return self._accounts.get(key)

    def remove_account(self, client_id: str, provider: str) -> None:
        key = (client_id, provider)
        with self._lock:
            self._accounts.pop(key, None)

    def list_accounts(self, client_id: str) -> dict[str, OAuthAccount]:
        with self._lock:
            return {
                provider: account
                for (saved_client_id, provider), account in self._accounts.items()
                if saved_client_id == client_id
            }

    def create_pending_state(self, provider: str, client_id: str, frontend_origin: str) -> str:
        self._cleanup_pending_states(max_age_seconds=900)
        state = secrets.token_urlsafe(32)
        pending = PendingOAuthState(
            provider=provider,
            client_id=client_id,
            frontend_origin=frontend_origin,
            created_at=time.time(),
        )
        with self._lock:
            self._pending_states[state] = pending
        return state

    def consume_pending_state(self, state: str) -> PendingOAuthState | None:
        with self._lock:
            return self._pending_states.pop(state, None)

    def _cleanup_pending_states(self, max_age_seconds: int) -> None:
        cutoff = time.time() - max_age_seconds
        with self._lock:
            expired = [key for key, value in self._pending_states.items() if value.created_at < cutoff]
            for key in expired:
                self._pending_states.pop(key, None)


oauth_store = OAuthStore()
