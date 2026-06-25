"""Google OAuth 2.0 authentication helper."""

import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from config import CREDENTIALS_FILE, TOKEN_FILE, LEGACY_TOKEN_FILE, SCOPES
from logger_setup import get_logger

log = get_logger("auth.google_auth")


class GoogleAuth:
    """Handles OAuth 2.0 flow and credential persistence."""

    def __init__(self):
        self.creds: Credentials | None = None
        self._load_existing_token()

    # ------------------------------------------------------------------
    def _load_existing_token(self):
        """Load saved token from disk if present and still valid."""
        for token_path in self._token_candidates():
            if not os.path.exists(token_path):
                continue
            try:
                self.creds = Credentials.from_authorized_user_file(token_path, SCOPES)
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                    self._save_token()
                    log.info("Refreshed expired Google token.")
                elif self.creds and self.creds.valid:
                    if token_path != TOKEN_FILE:
                        self._save_token()
                        log.info("Migrated legacy Google token to user data directory.")
                    log.info("Loaded valid Google token from disk.")
                else:
                    self.creds = None
                if self.creds:
                    return
            except Exception as exc:
                log.warning("Failed to load token from %s: %s", token_path, exc)
                self.creds = None

    # ------------------------------------------------------------------
    def _token_candidates(self) -> list[str]:
        candidates = [TOKEN_FILE]
        if LEGACY_TOKEN_FILE != TOKEN_FILE:
            candidates.append(LEGACY_TOKEN_FILE)
        return candidates

    # ------------------------------------------------------------------
    @property
    def is_connected(self) -> bool:
        return self.creds is not None and self.creds.valid

    # ------------------------------------------------------------------
    def authenticate(self) -> bool:
        """Run the OAuth installed-app flow (opens browser).
        Returns True on success."""
        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            self.creds = flow.run_local_server(
                port=8081,
                prompt="consent",
                success_message="Authentication successful! You can close this tab.",
            )
            self._save_token()
            log.info("Google OAuth completed successfully.")
            return True
        except Exception as exc:
            log.error("OAuth flow failed: %s", exc)
            return False

    # ------------------------------------------------------------------
    def _save_token(self):
        os.makedirs(os.path.dirname(TOKEN_FILE), exist_ok=True)
        temp_path = f"{TOKEN_FILE}.tmp"
        token_json = self.creds.to_json()
        with open(temp_path, "w", encoding="utf-8") as token_file:
            token_file.write(token_json)
        os.replace(temp_path, TOKEN_FILE)

    # ------------------------------------------------------------------
    def get_credentials(self) -> Credentials | None:
        """Return valid credentials, refreshing if needed."""
        if self.creds and self.creds.expired and self.creds.refresh_token:
            try:
                self.creds.refresh(Request())
                self._save_token()
            except Exception as exc:
                log.error("Token refresh failed: %s", exc)
                self.creds = None
        return self.creds
