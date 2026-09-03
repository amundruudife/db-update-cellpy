from types import SimpleNamespace

import pytest

import main
from src.deferred_cp.source_acquisition import SourceMetadata


def test_acquire_requires_an_account():
    with pytest.raises(SystemExit) as exc_info:
        main.parse_arguments(["acquire"])

    assert exc_info.value.code == 2


def test_acquire_cli_uses_keyring_and_stages_the_fixed_snapshot(monkeypatch, tmp_path, capsys):
    calls = []

    class FakeCredentialStore:
        def get_token(self, account):
            calls.append(("token", account))
            return "secret-token"

    class FakeTransport:
        def __init__(self):
            calls.append(("transport",))

    metadata = SourceMetadata(
        drive_id="drive-456",
        item_id="item-123",
        etag='"etag-789"',
        last_modified="2026-08-05T10:11:12Z",
        workbook_name="Cell_Log.xlsx",
        sheet_name="c&p",
        used_range="c&p!A1:S4590",
    )
    staged = tmp_path / "source_data" / "Cell_Log_CP.xlsx"

    def fake_acquire(resolver, session_factory, root):
        calls.append(("acquire", resolver, session_factory, root))
        return metadata, staged

    monkeypatch.setattr(main, "KeyringCredentialStore", FakeCredentialStore)
    monkeypatch.setattr(main, "GraphHttpTransport", FakeTransport)
    monkeypatch.setattr(main, "acquire_and_stage_snapshot", fake_acquire)

    result = main.main(["acquire", "--account", "ife-user", "--root", str(tmp_path)])

    captured = capsys.readouterr()
    assert result == 0
    assert calls[0] == ("token", "ife-user")
    assert calls[1] == ("transport",)
    assert calls[2][0] == "acquire"
    assert calls[2][3] == tmp_path
    assert "Cell_Log_CP.xlsx" in captured.out
    assert "secret-token" not in captured.out
    assert captured.err == ""
