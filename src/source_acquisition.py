"""Read-only acquisition boundaries for the hard-coded Cell Log source."""

import json
import base64
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote

from .contracts import SHAREPOINT_WORKBOOK_URL, SOURCE_SHEET_NAME


EXPECTED_WORKBOOK_NAME = "Cell_Log.xlsx"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
EXPECTED_USED_RANGE = re.compile(r"(?:c&p|'c&p')!A1:S[1-9][0-9]*$")


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

    @staticmethod
    def _share_token(url: str) -> str:
        encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
        return "u!" + encoded.rstrip("=")

    def resolve_drive_item(self, url: str):
        token = self._share_token(url)
        return self._get(f"/shares/{token}/driveItem")

    def read_used_range(self, identity: SourceIdentity, sheet_name: str):
        sheet = quote(sheet_name, safe="")
        return self._get(
            f"/drives/{identity.drive_id}/items/{identity.item_id}"
            f"/workbook/worksheets('{sheet}')/usedRange(valuesOnly=true)"
        )


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
    return metadata, values


def acquire_and_stage_snapshot(resolver, session_factory, root=None):
    """Acquire the source before staging a replacement local snapshot.

    Acquisition failures propagate without entering the snapshot replacement
    boundary.  This function has no production path or fallback source.
    """

    from .snapshot import write_snapshot

    metadata, values = acquire_source(resolver, session_factory)
    return metadata, write_snapshot(values, root=root)
