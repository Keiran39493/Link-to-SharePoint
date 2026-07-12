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
        return {"value": [{"id": "siteA", "displayName": "Corporate"}]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "root", "root")

    assert "followedSites" in calls[0]
    assert items == [{"name": "Corporate", "type": "folder", "driveId": "site:siteA", "itemId": "drives"}]


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
        return {"value": [{"id": "b!driveXYZ", "name": "Documents"}]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "site:siteA", "drives")

    assert calls[0] == f"{graph_client.GRAPH}/sites/siteA/drives?$select=id,name"
    assert items == [{"name": "Documents", "type": "folder", "driveId": "b!driveXYZ", "itemId": "root"}]


def test_drive_root_lists_children(monkeypatch):
    calls = []

    def fake_get(token, url):
        calls.append(url)
        return {"value": [
            {"name": "TAX", "folder": {"childCount": 4}, "id": "folder1"},
            {"name": "notes.txt", "size": 512},
        ]}

    monkeypatch.setattr(graph_client, "_get", fake_get)
    items = graph_client.list_children("tok", "b!driveXYZ", "root")

    assert calls[0] == f"{graph_client.GRAPH}/drives/b!driveXYZ/root/children"
    assert items[0] == {"name": "TAX", "type": "folder", "itemId": "folder1", "childCount": 4}
    assert items[1] == {"name": "notes.txt", "type": "file", "size": "512 B"}


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
