"""Thin Microsoft Graph REST wrapper for browsing + filing.

Browsing is addressed as an opaque (driveId, itemId) pair, same as the
/sp/list contract expects. Two pseudo-levels sit above real Graph drives so
a SharePoint *site* can be browsed before you know which document library
(drive) or folder you want:

  driveId="root", itemId="root"            -> list sites
                                               (Graph: /me/followedSites, or
                                               a fixed SITE_ALLOWLIST)
  driveId="site:<siteId>", itemId="drives" -> list that site's document
                                               libraries (Graph: /sites/{id}/drives)
  driveId=<real drive id>, itemId="root"
    or a real item id                      -> normal
                                               /drives/{driveId}/items/{itemId}/children

The task pane never needs to know about this -- it just round-trips
whatever driveId/itemId a listing handed it.
"""
from __future__ import annotations

import os
from urllib.parse import quote

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
UPLOAD_LIMIT_BYTES = 4 * 1024 * 1024  # Graph's simple-upload (PUT .../content) ceiling


class GraphError(Exception):
    pass


class GraphConflictError(GraphError):
    """Raised when the upload would silently overwrite an existing file
    (Graph returns 409 because we explicitly ask it to fail on conflict
    instead of replacing)."""

    def __init__(self, filename: str):
        self.filename = filename
        super().__init__(f"A file named '{filename}' already exists in this folder.")


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


def _get(token: str, url: str) -> dict:
    r = requests.get(url, headers=_headers(token), timeout=20)
    if not r.ok:
        raise GraphError(f"GET {url} -> {r.status_code}: {r.text[:300]}")
    return r.json()


def list_children(token: str, drive_id: str, item_id: str) -> list[dict]:
    if drive_id == "root" and item_id == "root":
        return _list_sites(token)
    if drive_id.startswith("site:") and item_id == "drives":
        return _list_drives(token, drive_id.split(":", 1)[1])
    anchor = "root" if item_id == "root" else f"items/{item_id}"
    data = _get(token, f"{GRAPH}/drives/{drive_id}/{anchor}/children")
    return [_shape_item(i) for i in data.get("value", [])]


def _list_sites(token: str) -> list[dict]:
    """Top level: sites the user can see. Defaults to /me/followedSites
    (only what they've followed); set SITE_ALLOWLIST (comma-separated
    Graph site IDs) to show a fixed set instead, e.g. if your org doesn't
    rely on people following sites."""
    allowlist = [s.strip() for s in os.environ.get("SITE_ALLOWLIST", "").split(",") if s.strip()]
    if allowlist:
        sites = [_get(token, f"{GRAPH}/sites/{site_id}?$select=id,displayName,webUrl") for site_id in allowlist]
    else:
        sites = _get(token, f"{GRAPH}/me/followedSites?$select=id,displayName,webUrl").get("value", [])
    return [
        {
            "name": s.get("displayName") or s.get("name") or "Untitled site",
            "type": "folder",
            "driveId": f"site:{s['id']}",
            "itemId": "drives",
            "webUrl": s.get("webUrl"),
        }
        for s in sites
    ]


def _list_drives(token: str, site_id: str) -> list[dict]:
    data = _get(token, f"{GRAPH}/sites/{site_id}/drives?$select=id,name,webUrl")
    return [
        {"name": d.get("name") or "Documents", "type": "folder", "driveId": d["id"], "itemId": "root", "webUrl": d.get("webUrl")}
        for d in data.get("value", [])
    ]


def _shape_item(item: dict) -> dict:
    if "folder" in item:
        return {
            "name": item["name"],
            "type": "folder",
            "itemId": item["id"],
            "childCount": item["folder"].get("childCount"),
            "webUrl": item.get("webUrl"),
        }
    size = item.get("size")
    return {
        "name": item["name"],
        "type": "file",
        "size": _human_size(size) if size is not None else None,
        "webUrl": item.get("webUrl"),
    }


def _human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def get_message_mime(token: str, message_id: str) -> bytes:
    """Fetches a message's raw MIME content (RFC 822 / .eml) via Graph's
    $value endpoint -- this already carries the full conversation context
    (quoted replies) that Outlook embedded in the message body."""
    url = f"{GRAPH}/me/messages/{quote(message_id, safe='')}/$value"
    r = requests.get(url, headers=_headers(token), timeout=30)
    if not r.ok:
        raise GraphError(f"GET {url} -> {r.status_code}: {r.text[:300]}")
    return r.content


def upload_file(token: str, drive_id: str, folder_item_id: str, filename: str, content: bytes) -> dict:
    if len(content) > UPLOAD_LIMIT_BYTES:
        raise GraphError(
            f"Attachment is {len(content)} bytes, over Graph's {UPLOAD_LIMIT_BYTES}-byte "
            "simple-upload limit. Larger files need a resumable upload session (not implemented here)."
        )
    # @microsoft.graph.conflictBehavior=fail tells Graph to reject the write
    # (409) instead of its default of silently replacing an existing file —
    # we never want an overwrite to happen invisibly; the caller is expected
    # to have the user pick a different name instead.
    if folder_item_id == "root":
        # Drive root: documented path syntax is /drives/{id}/root:/{name}:/content
        # (NOT items/root: — "root" is not a real parent item-id).
        url = f"{GRAPH}/drives/{drive_id}/root:/{quote(filename)}:/content?@microsoft.graph.conflictBehavior=fail"
    else:
        # A real folder item-id: address by parent-id path form.
        url = f"{GRAPH}/drives/{drive_id}/items/{folder_item_id}:/{quote(filename)}:/content?@microsoft.graph.conflictBehavior=fail"
    r = requests.put(
        url,
        headers={**_headers(token), "Content-Type": "application/octet-stream"},
        data=content,
        timeout=30,
    )
    if r.status_code == 409:
        raise GraphConflictError(filename)
    if not r.ok:
        raise GraphError(f"PUT {url} -> {r.status_code}: {r.text[:300]}")
    return r.json()
