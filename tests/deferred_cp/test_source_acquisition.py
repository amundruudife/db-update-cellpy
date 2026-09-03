import pytest

from src.deferred_cp.contracts import SHAREPOINT_WORKBOOK_URL
from src.deferred_cp.source_acquisition import (
    acquire_and_stage_snapshot,
    acquire_source,
    GraphWorkbookProvider,
    GraphRequestError,
    NonPersistentWorkbookSession,
    ReadOnlyGraphResolver,
    resolve_source_identity,
    resolve_source_metadata,
    SourceChangedError,
    write_source_manifest,
)


def test_sharepoint_url_is_the_token_free_canonical_identity():
    assert SHAREPOINT_WORKBOOK_URL == (
        "https://ifecloud.sharepoint.com/sites/UsersofIFEBatteryLab/"
        "General/00_Logs/Cell_Log.xlsx"
    )


def test_authentication_failure_preserves_accepted_snapshot_and_production(tmp_path):
    accepted_snapshot = tmp_path / "source_data" / "Cell_Log_CP.xlsx"
    accepted_snapshot.parent.mkdir()
    accepted_snapshot.write_bytes(b"accepted snapshot")
    production = tmp_path / "production.xlsx"
    production.write_bytes(b"production workbook")
    before_snapshot = accepted_snapshot.read_bytes()
    before_production = production.read_bytes()

    class FailingResolver:
        def resolve_drive_item(self, url):
            raise PermissionError("SharePoint authentication failed")

    with pytest.raises(PermissionError, match="authentication failed"):
        acquire_and_stage_snapshot(
            FailingResolver(),
            lambda metadata: pytest.fail("authentication failure must not open a session"),
            root=tmp_path,
        )

    assert accepted_snapshot.read_bytes() == before_snapshot
    assert production.read_bytes() == before_production


class FakeResolver:
    def __init__(self):
        self.urls = []

    def resolve_drive_item(self, url):
        self.urls.append(url)
        return {
            "id": "item-123",
            "parentReference": {"driveId": "drive-456"},
        }


def test_resolve_source_identity_uses_the_hard_coded_sharepoint_url():
    resolver = FakeResolver()

    identity = resolve_source_identity(resolver)

    assert resolver.urls == [SHAREPOINT_WORKBOOK_URL]
    assert identity.drive_id == "drive-456"
    assert identity.item_id == "item-123"


def test_resolve_source_identity_rejects_incomplete_drive_item():
    class IncompleteResolver:
        def resolve_drive_item(self, url):
            return {"id": "item-123"}

    with pytest.raises(ValueError, match="drive item identity"):
        resolve_source_identity(IncompleteResolver())


def test_resolve_source_metadata_records_item_and_used_range_metadata(tmp_path):
    class MetadataResolver:
        def resolve_drive_item(self, url):
            return {
                "id": "item-123",
                "parentReference": {"driveId": "drive-456"},
                "name": "Cell_Log.xlsx",
                "eTag": "\"etag-789\"",
                "lastModifiedDateTime": "2026-08-04T10:11:12Z",
            }

        def read_used_range(self, identity, sheet_name):
            assert identity.item_id == "item-123"
            assert sheet_name == "c&p"
            return {"sheetName": "c&p", "address": "c&p!A1:S4590"}

    metadata = resolve_source_metadata(MetadataResolver())
    manifest_path = write_source_manifest(metadata, tmp_path / "manifest.json")

    assert metadata.drive_id == "drive-456"
    assert metadata.item_id == "item-123"
    assert metadata.etag == '"etag-789"'
    assert metadata.last_modified == "2026-08-04T10:11:12Z"
    assert metadata.workbook_name == "Cell_Log.xlsx"
    assert metadata.sheet_name == "c&p"
    assert metadata.used_range == "c&p!A1:S4590"
    assert manifest_path.read_text(encoding="utf-8").find('"item_id": "item-123"') >= 0


def test_resolve_source_metadata_rejects_a_different_workbook_name():
    class WrongWorkbookResolver:
        def resolve_drive_item(self, url):
            return {
                "id": "item-123",
                "parentReference": {"driveId": "drive-456"},
                "name": "Other.xlsx",
                "eTag": "etag",
                "lastModifiedDateTime": "2026-08-04T10:11:12Z",
            }

        def read_used_range(self, identity, sheet_name):
            raise AssertionError("wrong workbook must stop before reading a sheet")

    with pytest.raises(ValueError, match="Unexpected source workbook"):
        resolve_source_metadata(WrongWorkbookResolver())


def test_resolve_source_metadata_rejects_a_used_range_from_another_sheet():
    class WrongSheetResolver:
        def resolve_drive_item(self, url):
            return {
                "id": "item-123",
                "parentReference": {"driveId": "drive-456"},
                "name": "Cell_Log.xlsx",
                "eTag": "etag",
                "lastModifiedDateTime": "2026-08-04T10:11:12Z",
            }

        def read_used_range(self, identity, sheet_name):
            return {"sheetName": "c&p", "address": "log!A1:S4590"}

    with pytest.raises(ValueError, match="used-range address"):
        resolve_source_metadata(WrongSheetResolver())


def test_authentication_or_network_failure_stops_without_fallback():
    session_calls = []

    class FailingResolver:
        def resolve_drive_item(self, url):
            raise ConnectionError("SharePoint unavailable")

    def session_factory(metadata):
        session_calls.append(metadata)
        raise AssertionError("a failed source resolution must not open a session")

    with pytest.raises(ConnectionError, match="SharePoint unavailable"):
        acquire_source(FailingResolver(), session_factory)

    assert session_calls == []


def test_graph_resolver_exposes_only_read_requests():
    class ReadOnlyTransport:
        def __init__(self):
            self.calls = []

        def get(self, url, headers):
            self.calls.append((url, headers))
            if "/sites/ifecloud.sharepoint.com:/" in url:
                return {"id": "site-789"}
            if "/drive/root:/" in url:
                return {
                    "id": "item-123",
                    "parentReference": {"driveId": "drive-456"},
                }
            return {"sheetName": "c&p", "address": "c&p!A1:S4590"}

    transport = ReadOnlyTransport()
    resolver = ReadOnlyGraphResolver(transport, access_token="token")

    identity = resolve_source_identity(resolver)
    used_range = resolver.read_used_range(identity, "c&p")

    assert identity.item_id == "item-123"
    assert used_range["address"] == "c&p!A1:S4590"
    assert len(transport.calls) == 3
    assert transport.calls[0][0].endswith(
        "/sites/ifecloud.sharepoint.com:/sites/UsersofIFEBatteryLab"
    )
    assert transport.calls[1][0].endswith(
        "/sites/site-789/drive/root:/General/00_Logs/Cell_Log.xlsx"
    )
    assert all(call[1]["Authorization"] == "Bearer token" for call in transport.calls)
    shares_route = "/" + "shares" + "/"
    assert not any(shares_route in call[0] for call in transport.calls)
    assert not any(hasattr(resolver, method) for method in ("post", "put", "patch", "delete"))


def test_workbook_session_is_non_persistent_and_closes_after_read():
    class SessionProvider:
        def __init__(self):
            self.events = []

        def create_session(self, drive_id, item_id, persist_changes):
            self.events.append(("create", drive_id, item_id, persist_changes))
            return "session-123"

        def read_values(self, session_id, sheet_name, used_range):
            self.events.append(("read", session_id, sheet_name, used_range))
            return [[1, "ok"]]

        def close_session(self, session_id):
            self.events.append(("close", session_id))

    provider = SessionProvider()
    session = NonPersistentWorkbookSession(
        provider,
        drive_id="drive-456",
        item_id="item-123",
    )

    with session as active:
        assert active.read_values("c&p", "c&p!A1:S2") == [[1, "ok"]]

    assert provider.events == [
        ("create", "drive-456", "item-123", False),
        ("read", "session-123", "c&p", "c&p!A1:S2"),
        ("close", "session-123"),
    ]


def test_graph_workbook_provider_uses_non_persistent_read_session():
    class GraphTransport:
        def __init__(self):
            self.calls = []

        def post(self, url, headers=None, json_body=None):
            self.calls.append(("post", url, headers, json_body))
            return {"id": "session-123"}

        def get(self, url, headers=None):
            self.calls.append(("get", url, headers))
            return {"values": [[101, "ok"]]}

    transport = GraphTransport()
    provider = GraphWorkbookProvider(transport, access_token="token")
    session = NonPersistentWorkbookSession(provider, "drive-456", "item-123")

    with session as active:
        assert active.read_values("c&p", "c&p!A1:S2") == [[101, "ok"]]

    assert transport.calls[0][0] == "post"
    assert transport.calls[0][3] == {"persistChanges": False}
    assert transport.calls[0][2]["Authorization"] == "Bearer token"
    assert transport.calls[1][0] == "get"
    assert "range(address='A1%3AS2')" in transport.calls[1][1]
    assert transport.calls[1][2]["workbook-session-id"] == "session-123"
    assert transport.calls[2][0] == "post"
    assert transport.calls[2][2]["workbook-session-id"] == "session-123"


class _VersionedResolver:
    def __init__(self, etags):
        self.etags = iter(etags)
        self.current_etag = None

    def resolve_drive_item(self, url):
        self.current_etag = next(self.etags, self.current_etag)
        return {
            "id": "item-123",
            "parentReference": {"driveId": "drive-456"},
            "name": "Cell_Log.xlsx",
            "eTag": self.current_etag,
            "lastModifiedDateTime": "2026-08-04T10:11:12Z",
        }

    def read_used_range(self, identity, sheet_name):
        return {"sheetName": "c&p", "address": "c&p!A1:S4590"}


class _ReadSession:
    def __init__(self, values):
        self.values = values
        self.closed = False

    def __enter__(self):
        return self

    def read_values(self, sheet_name, used_range):
        return self.values

    def __exit__(self, exc_type, exc_value, traceback):
        self.closed = True
        return False


def test_acquire_source_rejects_a_changed_etag_after_read():
    resolver = _VersionedResolver(['"etag-1"', '"etag-2"'])
    session = _ReadSession([["evaluated"]])

    with pytest.raises(SourceChangedError, match="eTag"):
        acquire_source(resolver, lambda metadata: session)

    assert session.closed


def test_acquire_source_returns_values_when_source_metadata_is_unchanged():
    resolver = _VersionedResolver(['"etag-1"', '"etag-1"'])
    values = [[None] * 19 for _ in range(4590)]
    session = _ReadSession(values)

    metadata, returned_values = acquire_source(resolver, lambda metadata: session)

    assert metadata.etag == '"etag-1"'
    assert returned_values == values
    assert session.closed


@pytest.mark.parametrize("row_count", [4589, 4591], ids=["one-row-short", "one-row-extra"])
def test_acquire_source_rejects_values_with_wrong_terminal_row_count(row_count):
    resolver = _VersionedResolver(['"etag-1"', '"etag-1"'])
    session = _ReadSession([[None] * 19 for _ in range(row_count)])

    with pytest.raises(ValueError, match="expected 4590"):
        acquire_source(resolver, lambda metadata: session)

    assert session.closed


@pytest.mark.parametrize(
    "address",
    ["c&p!A1:S", "c&p!A1:R2", "c&p!A1:S0"],
    ids=["missing-terminal-row", "wrong-terminal-column", "zero-terminal-row"],
)
def test_acquire_source_rejects_malformed_terminal_range(address):
    class Resolver(_VersionedResolver):
        def read_used_range(self, identity, sheet_name):
            return {"sheetName": "c&p", "address": address}

    session = _ReadSession([[None] * 19])
    with pytest.raises(ValueError, match="used-range address"):
        acquire_source(Resolver(['"etag-1"', '"etag-1"']), lambda metadata: session)
    assert session.closed is False


def test_acquire_source_rejects_malformed_terminal_row_and_closes_session():
    resolver = _VersionedResolver(['"etag-1"', '"etag-1"'])
    session = _ReadSession([[None] * 18 for _ in range(4590)])

    with pytest.raises(ValueError, match="malformed c&p rows"):
        acquire_source(resolver, lambda metadata: session)

    assert session.closed


def test_acquire_and_stage_snapshot_preserves_snapshot_and_cleans_cardinality_temp(tmp_path):
    accepted = tmp_path / "source_data" / "Cell_Log_CP.xlsx"
    accepted.parent.mkdir()
    accepted.write_bytes(b"accepted")
    resolver = _VersionedResolver(['"etag-1"', '"etag-1"'])
    session = _ReadSession([[None] * 19])

    with pytest.raises(ValueError, match="expected 4590"):
        acquire_and_stage_snapshot(resolver, lambda metadata: session, root=tmp_path)

    assert session.closed
    assert accepted.read_bytes() == b"accepted"
    assert not list(accepted.parent.glob(".Cell_Log_CP.*"))


def test_graph_resolver_fails_closed_for_missing_site_or_item_identity():
    class MissingSite:
        def get(self, url, headers=None):
            return {}

    with pytest.raises(GraphRequestError, match="site identity"):
        ReadOnlyGraphResolver(MissingSite(), "token").resolve_drive_item(SHAREPOINT_WORKBOOK_URL)

    class MissingItem:
        def get(self, url, headers=None):
            if "/drive/root:/" in url:
                return {"id": "item"}
            return {"id": "site"}

    with pytest.raises(ValueError, match="drive item identity"):
        resolve_source_identity(ReadOnlyGraphResolver(MissingItem(), "token"))
