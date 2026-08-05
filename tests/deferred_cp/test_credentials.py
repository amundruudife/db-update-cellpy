import pytest

from src.deferred_cp.credentials import CredentialStoreError, KeyringCredentialStore


class FakeKeyring:
    def __init__(self):
        self.values = {}
        self.calls = []

    def get_password(self, service, account):
        self.calls.append(("get", service, account))
        return self.values.get((service, account))

    def set_password(self, service, account, token):
        self.calls.append(("set", service, account))
        self.values[(service, account)] = token


def test_keyring_store_round_trips_token_without_file_or_environment_fallback():
    backend = FakeKeyring()
    store = KeyringCredentialStore(backend=backend, service="cell-log-updater")

    store.set_token("ife-user", "secret-token")

    assert store.get_token("ife-user") == "secret-token"
    assert backend.calls == [
        ("set", "cell-log-updater", "ife-user"),
        ("get", "cell-log-updater", "ife-user"),
    ]


def test_keyring_store_rejects_missing_token():
    store = KeyringCredentialStore(backend=FakeKeyring())

    with pytest.raises(CredentialStoreError, match="No access token"):
        store.get_token("ife-user")


def test_keyring_store_rejects_empty_tokens():
    store = KeyringCredentialStore(backend=FakeKeyring())

    with pytest.raises(ValueError, match="non-empty"):
        store.set_token("ife-user", "")
