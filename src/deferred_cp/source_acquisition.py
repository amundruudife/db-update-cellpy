"""Read-only acquisition boundaries for the hard-coded Cell Log source."""

import json
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from .contracts import (
    SHAREPOINT_DEFAULT_DRIVE_ITEM_PATH,
    SHAREPOINT_HOSTNAME,
    SHAREPOINT_SITE_PATH,
    SHAREPOINT_WORKBOOK_URL,
    SOURCE_COLUMN_COUNT,
    SOURCE_SHEET_NAME,
)


EXPECTED_WORKBOOK_NAME = "Cell_Log.xlsx"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
EXPECTED_USED_RANGE = re.compile(r"(?:c&p|'c&p')!A1:S([1-9][0-9]*)$")


@dataclass(frozen=True)
class SourceIdentity:
    """Immutable SharePoint identity used for a source retrieval."""

    drive_id: str
    item_id: str


@dataclass(frozen=True)
class SourceMetadata:
    """Source identity and retrieval metadata retained with a snapshot."""

    drive_id: str
    item_id: str
    etag: str
    last_modified: str
    workbook_name: str
    sheet_name: str
    used_range: str
    source_url: str = SHAREPOINT_WORKBOOK_URL

    def as_dict(self) -> dict:
        return {
            "source_url": self.source_url,
            "drive_id": self.drive_id,
            "item_id": self.item_id,
            "etag": self.etag,
            "last_modified": self.last_modified,
            "workbook_name": self.workbook_name,
            "sheet_name": self.sheet_name,
            "used_range": self.used_range,
        }


class SourceChangedError(RuntimeError):
    """Raised when the source changes during one acquisition."""


class GraphRequestError(RuntimeError):
    """Raised when Microsoft Graph cannot complete a source request."""


class GraphHttpTransport:
    """Small JSON transport for the read-only source acquisition flow.

    Authentication is supplied by the caller in request headers. This class
    does not expose a source-write operation; ``POST`` is used only by the
    workbook session provider for non-persistent session lifecycle calls.
    """

    def get(self, url: str, headers=None):
        return self._request("GET", url, headers=headers)

    def post(self, url: str, headers=None, json_body=None):
        return self._request("POST", url, headers=headers, json_body=json_body)

    @staticmethod
    def _request(method: str, url: str, *, headers=None, json_body=None):
        request_headers = {"Accept": "application/json", **(headers or {})}
        body = None
        if json_body is not None:
            request_headers.setdefault("Content-Type", "application/json")
            body = json.dumps(json_body).encode("utf-8")
        request = Request(
            url,
            data=body,
            headers=request_headers,
            method=method,
        )
        try:
            with urlopen(request, timeout=60) as response:
                payload = response.read()
        except HTTPError as exc:
            raise GraphRequestError(
                f"Graph {method} request failed with HTTP {exc.code}"
            ) from exc
        except URLError as exc:
            raise GraphRequestError(f"Graph {method} request failed: {exc.reason}") from exc

        if not payload:
            return {}
        try:
            result = json.loads(payload.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise GraphRequestError(f"Graph {method} request returned invalid JSON") from exc
        if not isinstance(result, dict):
            raise GraphRequestError(f"Graph {method} request returned a non-object response")
        return result


def ensure_source_unchanged(before: SourceMetadata, after: SourceMetadata) -> None:
    """Require source identity, range, version, and freshness to match."""

    fields = (
        ("drive item", "drive_id"),
        ("drive item", "item_id"),
        ("workbook", "workbook_name"),
        ("sheet", "sheet_name"),
        ("used-range address", "used_range"),
        ("eTag", "etag"),
        ("last-modified timestamp", "last_modified"),
    )
    changed = [label for label, field in fields if getattr(before, field) != getattr(after, field)]
    if changed:
        raise SourceChangedError(
            "Source metadata changed during acquisition: " + ", ".join(changed)
        )


class ReadOnlyGraphResolver:
    """Resolve source metadata using a transport that only supports GET."""

    def __init__(self, transport, access_token: str, graph_api_base: str = GRAPH_API_BASE):
        if not access_token:
            raise ValueError("An access token is required")
        self._transport = transport
        self._access_token = access_token
        self._graph_api_base = graph_api_base.rstrip("/")

    def _get(self, path: str):
        return self._transport.get(
            f"{self._graph_api_base}{path}",
            headers={"Authorization": f"Bearer {self._access_token}"},
        )

    def resolve_drive_item(self, url: str):
        if url != SHAREPOINT_WORKBOOK_URL:
            raise ValueError("source resolution is restricted to the approved workbook")
        site_path = quote(SHAREPOINT_SITE_PATH, safe="/")
        site = self._get(f"/sites/{SHAREPOINT_HOSTNAME}:/{site_path}")
        site_id = site.get("id") if isinstance(site, dict) else None
        if not site_id:
            raise GraphRequestError("Graph site lookup lacks a site identity")
        item_path = quote(SHAREPOINT_DEFAULT_DRIVE_ITEM_PATH, safe="/")
        item = self._get(f"/sites/{quote(str(site_id), safe='')}/drive/root:/{item_path}")
        if not isinstance(item, dict):
            raise GraphRequestError("Graph default-drive item lookup returned an invalid response")
        return item

    def read_used_range(self, identity: SourceIdentity, sheet_name: str):
        sheet = quote(sheet_name, safe="")
        return self._get(
            f"/drives/{identity.drive_id}/items/{identity.item_id}"
            f"/workbook/worksheets('{sheet}')/usedRange(valuesOnly=true)"
        )


class GraphWorkbookProvider:
    """Read evaluated values through a non-persistent Graph workbook session."""

    def __init__(self, transport, access_token: str, graph_api_base: str = GRAPH_API_BASE):
        if not access_token:
            raise ValueError("An access token is required")
        self._transport = transport
        self._access_token = access_token
        self._graph_api_base = graph_api_base.rstrip("/")

    def _workbook_path(self, drive_id: str, item_id: str) -> str:
        return (
            f"{self._graph_api_base}/drives/{quote(str(drive_id), safe='')}"
            f"/items/{quote(str(item_id), safe='')}/workbook"
        )

    def _headers(self, session_id=None):
        headers = {"Authorization": f"Bearer {self._access_token}"}
        if session_id is not None:
            headers["workbook-session-id"] = session_id
        return headers

    def create_session(self, drive_id: str, item_id: str, persist_changes: bool):
        if persist_changes is not False:
            raise ValueError("Source acquisition requires a non-persistent workbook session")
        self._drive_id = str(drive_id)
        self._item_id = str(item_id)
        response = self._transport.post(
            f"{self._workbook_path(drive_id, item_id)}/createSession",
            headers=self._headers(),
            json_body={"persistChanges": False},
        )
        session_id = response.get("id") if isinstance(response, dict) else None
        if not session_id:
            raise GraphRequestError("Graph createSession response lacks a session id")
        return str(session_id)

    def read_values(self, session_id: str, sheet_name: str, used_range: str):
        if sheet_name != SOURCE_SHEET_NAME:
            raise ValueError("Source acquisition is restricted to the c&p sheet")
        if not EXPECTED_USED_RANGE.fullmatch(str(used_range)):
            raise ValueError("Source acquisition is restricted to the approved c&p A1:S range")
        address = str(used_range).split("!", 1)[-1]
        encoded_sheet = quote(sheet_name, safe="")
        encoded_address = quote(address, safe="")
        response = self._transport.get(
            f"{self._graph_api_base}/drives/{quote(str(self._drive_id), safe='')}"
            f"/items/{quote(str(self._item_id), safe='')}/workbook"
            f"/worksheets('{encoded_sheet}')/range(address='{encoded_address}')",
            headers=self._headers(session_id),
        )
        values = response.get("values") if isinstance(response, dict) else None
        if not isinstance(values, list):
            raise GraphRequestError("Graph range response lacks evaluated values")
        return values

    def close_session(self, session_id: str):
        self._transport.post(
            f"{self._workbook_path(self._drive_id, self._item_id)}/closeSession",
            headers=self._headers(session_id),
        )

    def bind_workbook(self, drive_id: str, item_id: str):
        """Bind the provider to one resolved source item for session reads."""

        self._drive_id = str(drive_id)
        self._item_id = str(item_id)
        return self


class NonPersistentWorkbookSession:
    """Context-managed workbook session that cannot persist source changes."""

    def __init__(self, provider, drive_id: str, item_id: str):
        self._provider = provider
        self._drive_id = drive_id
        self._item_id = item_id
        self._session_id = None

    def __enter__(self):
        self._session_id = self._provider.create_session(
            self._drive_id,
            self._item_id,
            persist_changes=False,
        )
        return self

    def read_values(self, sheet_name: str, used_range: str):
        if self._session_id is None:
            raise RuntimeError("Workbook session is not open")
        return self._provider.read_values(self._session_id, sheet_name, used_range)

    def __exit__(self, exc_type, exc_value, traceback):
        if self._session_id is not None:
            session_id = self._session_id
            self._session_id = None
            self._provider.close_session(session_id)
        return False


def _identity_from_item(item) -> SourceIdentity:
    try:
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
    except (KeyError, TypeError) as exc:
        raise ValueError("Resolved response lacks a drive item identity") from exc

    if not drive_id or not item_id:
        raise ValueError("Resolved response lacks a drive item identity")

    return SourceIdentity(drive_id=str(drive_id), item_id=str(item_id))


def resolve_source_identity(resolver) -> SourceIdentity:
    """Resolve the approved URL to a drive and item identity.

    ``resolver`` is deliberately injected so the contract can be tested
    without credentials or network access. No alternate URL or local source
    is considered when resolution fails.
    """

    return _identity_from_item(resolver.resolve_drive_item(SHAREPOINT_WORKBOOK_URL))


def resolve_source_metadata(resolver) -> SourceMetadata:
    """Resolve and validate the source metadata needed for one retrieval."""

    item = resolver.resolve_drive_item(SHAREPOINT_WORKBOOK_URL)
    identity = _identity_from_item(item)
    workbook_name = item.get("name") if isinstance(item, dict) else None
    etag = item.get("eTag") if isinstance(item, dict) else None
    last_modified = item.get("lastModifiedDateTime") if isinstance(item, dict) else None
    if workbook_name != EXPECTED_WORKBOOK_NAME:
        raise ValueError(f"Unexpected source workbook: {workbook_name!r}")
    if not etag or not last_modified:
        raise ValueError("Source metadata lacks eTag or last-modified timestamp")

    used_range = resolver.read_used_range(identity, SOURCE_SHEET_NAME)
    if not isinstance(used_range, dict):
        raise ValueError("Source used-range response is invalid")
    if used_range.get("sheetName") != SOURCE_SHEET_NAME:
        raise ValueError("Source sheet identity is not c&p")
    address = used_range.get("address")
    if not address or not EXPECTED_USED_RANGE.fullmatch(str(address)):
        raise ValueError("Source used-range address is missing")

    return SourceMetadata(
        drive_id=identity.drive_id,
        item_id=identity.item_id,
        etag=str(etag),
        last_modified=str(last_modified),
        workbook_name=workbook_name,
        sheet_name=SOURCE_SHEET_NAME,
        used_range=str(address),
    )


def write_source_manifest(metadata: SourceMetadata, path) -> Path:
    """Write the metadata record used to identify an accepted source read."""

    manifest_path = Path(path)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(
        json.dumps(metadata.as_dict(), indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return manifest_path


def acquire_source(resolver, session_factory):
    """Read the approved source through one injected workbook session.

    Resolution and reads intentionally have no fallback path. Any resolver,
    authentication, or network exception aborts the acquisition.
    """

    metadata = resolve_source_metadata(resolver)
    with session_factory(metadata) as session:
        values = session.read_values(metadata.sheet_name, metadata.used_range)
    after = resolve_source_metadata(resolver)
    ensure_source_unchanged(metadata, after)
    match = EXPECTED_USED_RANGE.fullmatch(metadata.used_range)
    if match is None:
        raise ValueError("Source used-range address is malformed")
    expected_rows = int(match.group(1))
    if not isinstance(values, list):
        raise ValueError("Graph range values must be an array")
    if len(values) != expected_rows:
        raise ValueError(
            f"Graph range returned {len(values)} rows; expected {expected_rows}"
        )
    if any(
        not isinstance(row, (list, tuple)) or len(row) != SOURCE_COLUMN_COUNT
        for row in values
    ):
        raise ValueError("Graph range returned malformed c&p rows")
    return metadata, values


def acquire_and_stage_snapshot(resolver, session_factory, root=None):
    """Acquire the source before staging a replacement local snapshot.

    Acquisition failures propagate without entering the snapshot replacement
    boundary.  This function has no production path or fallback source.
    """

    from .snapshot import write_snapshot

    metadata, values = acquire_source(resolver, session_factory)
    return metadata, write_snapshot(values, root=root)
