"""Google Service Account authentication helper."""

import json
import os

from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request

from config import SERVICE_ACCOUNT_FILE, SCOPES
from logger_setup import get_logger

log = get_logger("auth.service_account")


class ServiceAccountAuth:
    """Handles Service Account credential loading and refresh."""

    def __init__(self):
        self.creds: Credentials | None = None
        self._sa_email: str = ""
        self._load_credentials()

    def _load_credentials(self):
        """Load service account credentials from JSON file."""
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            log.info("Service account file not found: %s", SERVICE_ACCOUNT_FILE)
            return
        try:
            self.creds = Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES
            )
            # Extract email from the JSON file
            with open(SERVICE_ACCOUNT_FILE, "r", encoding="utf-8") as f:
                sa_data = json.load(f)
            self._sa_email = sa_data.get("client_email", "")
            log.info("Service account loaded: %s", self._sa_email)
        except Exception as exc:
            log.error("Failed to load service account: %s", exc)
            self.creds = None
            self._sa_email = ""

    @property
    def is_configured(self) -> bool:
        """True if service account JSON file exists."""
        return os.path.exists(SERVICE_ACCOUNT_FILE)

    @property
    def is_ready(self) -> bool:
        """True if credentials are loaded and usable."""
        return self.creds is not None

    @property
    def service_account_email(self) -> str:
        """Return the service account email address."""
        return self._sa_email

    def get_credentials(self) -> Credentials | None:
        """Return valid credentials, refreshing if needed."""
        if self.creds is None:
            self._load_credentials()
        if self.creds is None:
            return None
        try:
            if self.creds.expired:
                self.creds.refresh(Request())
                log.debug("Service account token refreshed.")
        except Exception as exc:
            log.error("Token refresh failed: %s", exc)
            self.creds = None
        return self.creds

    def reload(self):
        """Force reload credentials from disk."""
        self.creds = None
        self._sa_email = ""
        self._load_credentials()
