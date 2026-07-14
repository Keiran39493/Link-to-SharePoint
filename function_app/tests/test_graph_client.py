import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import graph_client  # noqa: E402


def test_human_size_formatting():
    assert graph_client._human_size(0) == "0 B"
    assert graph_client._human_size(999) == "999 B"
    assert graph_client._human_size(2048) == "2.0 KB"
    assert graph_client._human_size(5 * 1024 * 1024) == "5.0 MB"


def test_root_lists_sites(monkeypatch):
    calls = []

    def fake_get(token, url):
        calls.append(url)
        return {"value": [{"id": "siteA", "displayName": "Corporate", "webUrl": "https://contoso.sharepoint.com/sites/Corporate"}]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "root", "root")

    assert "followedSites" in calls[0]
    assert "webUrl" in calls[0]
    assert items == [{
        "name": "Corporate", "type": "folder", "driveId": "site:siteA", "itemId": "drives",
        "webUrl": "https://contoso.sharepoint.com/sites/Corporate",
    }]


def test_root_uses_site_allowlist_when_set(monkeypatch):
    monkeypatch.setenv("SITE_ALLOWLIST", "siteA, siteB")
    calls = []

    def fake_get(token, url):
        calls.append(url)
        site_id = "siteA" if "siteA" in url else "siteB"
        return {"id": site_id, "displayName": site_id.upper()}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "root", "root")

    assert len(calls) == 2
    assert {i["driveId"] for i in items} == {"site:siteA", "site:siteB"}
    monkeypatch.delenv("SITE_ALLOWLIST", raising=False)


def test_site_pseudo_level_lists_drives(monkeypatch):
    calls = []

    def fake_get(token, url):
        calls.append(url)
        return {"value": [{"id": "b!driveXYZ", "name": "Documents", "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents"}]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "site:siteA", "drives")

    assert calls[0] == f"{graph_client.GRAPH}/sites/siteA/drives?$select=id,name,webUrl"
    assert items == [{
        "name": "Documents", "type": "folder", "driveId": "b!driveXYZ", "itemId": "root",
        "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents",
    }]


def test_drive_root_lists_children(monkeypatch):
    calls = []

    def fake_get(token, url):
        calls.append(url)
        return {"value": [
            {"name": "TAX", "folder": {"childCount": 4}, "id": "folder1", "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents/TAX"},
            {"name": "notes.txt", "size": 512, "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents/notes.txt"},
        ]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "b!driveXYZ", "root")

    assert calls[0] == f"{graph_client.GRAPH}/drives/b!driveXYZ/root/children"
    assert items[0] == {
        "name": "TAX", "type": "folder", "itemId": "folder1", "childCount": 4,
        "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents/TAX",
    }
    assert items[1] == {
        "name": "notes.txt", "type": "file", "size": "512 B",
        "webUrl": "https://contoso.sharepoint.com/sites/Corporate/Documents/notes.txt",
    }


def test_drive_subfolder_lists_children_by_item_id(monkeypatch):
    calls = []

    def fake_get(token, url):
        calls.append(url)
        return {"value": []}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    graph_client.list_children("tok", "b!driveXYZ", "folder1")

    assert calls[0] == f"{graph_client.GRAPH}/drives/b!driveXYZ/items/folder1/children"


def test_upload_rejects_oversized_content():
    try:
        graph_client.upload_file("tok", "d", "root", "big.bin", b"x" * (graph_client.UPLOAD_LIMIT_BYTES + 1))
        assert False, "expected GraphError"
    except graph_client.GraphError as e:
        assert "simple-upload" in str(e)


def test_upload_requests_fail_on_conflict(monkeypatch):
    calls = []

    class FakeResponse:
        ok = True
        status_code = 201

        def json(self):
            return {"id": "new-item"}

    def fake_put(url, headers=None, data=None, timeout=None):
        calls.append(url)
        return FakeResponse()

    monkeypatch.setattr(graph_client.requests, "put", fake_put)
    graph_client.upload_file("tok", "b!driveXYZ", "folder1", "notes.txt", b"hello")

    assert "@microsoft.graph.conflictBehavior=fail" in calls[0]


def test_upload_raises_conflict_error_on_409(monkeypatch):
    class FakeResponse:
        ok = False
        status_code = 409
        text = "nameAlreadyExists"

    monkeypatch.setattr(graph_client.requests, "put", lambda url, headers=None, data=None, timeout=None: FakeResponse())

    try:
        graph_client.upload_file("tok", "b!driveXYZ", "folder1", "notes.txt", b"hello")
        assert False, "expected GraphConflictError"
    except graph_client.GraphConflictError as e:
        assert e.filename == "notes.txt"
        assert "already exists" in str(e)


def test_get_message_mime_returns_raw_bytes(monkeypatch):
    calls = []

    class FakeResponse:
        ok = True
        content = b"From: a@b.com\r\nSubject: Hi\r\n\r\nBody"

    def fake_get(url, headers=None, timeout=None):
        calls.append((url, headers))
        return FakeResponse()

    monkeypatch.setattr(graph_client.requests, "get", fake_get)
    content = graph_client.get_message_mime("tok", "AAMkAGI+id/with=chars")

    assert content == b"From: a@b.com\r\nSubject: Hi\r\n\r\nBody"
    assert calls[0][0] == f"{graph_client.GRAPH}/me/messages/AAMkAGI%2Bid%2Fwith%3Dchars/$value"
    assert calls[0][1]["Authorization"] == "Bearer tok"


def test_get_message_mime_raises_on_error(monkeypatch):
    class FakeResponse:
        ok = False
        status_code = 404
        text = "not found"

    monkeypatch.setattr(graph_client.requests, "get", lambda url, headers=None, timeout=None: FakeResponse())

    try:
        graph_client.get_message_mime("tok", "missing-id")
        assert False, "expected GraphError"
    except graph_client.GraphError as e:
        assert "404" in str(e)
