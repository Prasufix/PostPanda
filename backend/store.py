from __future__ import annotations

from threading import Lock
from uuid import uuid4

import pandas as pd


class SessionStore:
    def __init__(self) -> None:
        self._items: dict[str, pd.DataFrame] = {}
        self._lock = Lock()

    def create(self, dataframe: pd.DataFrame) -> str:
        session_id = uuid4().hex
        with self._lock:
            self._items[session_id] = dataframe
        return session_id

    def get(self, session_id: str) -> pd.DataFrame | None:
        with self._lock:
            return self._items.get(session_id)


store = SessionStore()
