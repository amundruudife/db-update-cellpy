"""OS/user credential storage for source acquisition."""


DEFAULT_SERVICE = "ife-cell-log-updater"


class CredentialStoreError(RuntimeError):
    """Raised when the approved credential store cannot provide a token."""


class KeyringCredentialStore:
    """Store Graph tokens through the user's OS-backed keyring.

    The import is lazy so validation and tests do not require a credential
    package. There is intentionally no file, environment, or prompt fallback.
    """

    def __init__(self, backend=None, service: str = DEFAULT_SERVICE):
        if backend is None:
            try:
                import keyring as backend
            except ImportError as exc:
                raise CredentialStoreError(
                    "The OS credential store requires the optional keyring package"
                ) from exc
        self._backend = backend
        self._service = service

    def get_token(self, account: str) -> str:
        token = self._backend.get_password(self._service, account)
        if not token:
            raise CredentialStoreError(f"No access token stored for account {account!r}")
        return token

    def set_token(self, account: str, token: str) -> None:
        if not token:
            raise ValueError("Access token must be non-empty")
        self._backend.set_password(self._service, account, token)
