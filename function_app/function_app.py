"""The three endpoints the task pane calls (see ../Setup for the contracts).

Auth model: the incoming request carries a Bearer token minted for THIS
backend's exposed scope (Nested App Authentication in the add-in). Every
route validates that token, then -- where a Graph call is needed --
exchanges it on-behalf-of the user for a Graph-scoped token, so Graph only
ever sees what the signed-in user can access.
"""
from __future__ import annotations

import base64
import json
import logging
import os

import azure.functions as func

import graph_client
import recommender
import storage
from auth import AuthError, validate_incoming_token, get_graph_token_on_behalf_of

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

logging.warning(">>> FUNCTION_APP LOADED - auth debug logging enabled <<<")

GRAPH_SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All"]


def _cors_headers() -> dict:
    return {
        "Access-Control-Allow-Origin": os.environ.get("ALLOWED_ORIGIN", "*"),
        "Access-Control-Allow-Headers": "authorization,content-type",
        "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    }


def _json(payload: dict, status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(payload), status_code=status, mimetype="application/json", headers=_cors_headers()
    )


def _preflight() -> func.HttpResponse:
    return func.HttpResponse(status_code=204, headers=_cors_headers())


def _graph_token(req: func.HttpRequest) -> str:
    """Validates the caller, then OBO-exchanges for a Graph token. Raises
    AuthError (caller should turn that into a 401) on any failure."""
    return get_graph_token_on_behalf_of(req.headers.get("Authorization"), GRAPH_SCOPES)


@app.route(route="sp/list", methods=["GET", "OPTIONS"])
def list_children(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return _preflight()
    try:
        graph_token = _graph_token(req)
    except AuthError as e:
        logging.warning("AUTH FAIL [%s]: %s", req.url, e)
        return _json({"error": str(e)}, 401)

    drive_id = req.params.get("driveId", "root")
    item_id = req.params.get("itemId", "root")
    try:
        items = graph_client.list_children(graph_token, drive_id, item_id)
    except graph_client.GraphError as e:
        logging.warning("listChildren: %s", e)
        return _json({"error": str(e)}, 502)
    return _json({"items": items})


@app.route(route="sp/recommend", methods=["POST", "OPTIONS"])
def recommend(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return _preflight()
    try:
        validate_incoming_token(req.headers.get("Authorization"))
    except AuthError as e:
        logging.warning("AUTH FAIL [%s]: %s", req.url, e)
        return _json({"error": str(e)}, 401)

    try:
        body = req.get_json()
    except ValueError:
        return _json({"error": "Expected a JSON body"}, 400)

    document = body.get("document", {})
    history = storage.fetch_recent_decisions()
    recs = recommender.recommend(document, history)
    return _json({"recommendations": recs})


@app.route(route="sp/file", methods=["POST", "OPTIONS"])
def record_choice(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return _preflight()
    try:
        graph_token = _graph_token(req)
    except AuthError as e:
        logging.warning("AUTH FAIL [%s]: %s", req.url, e)
        return _json({"error": str(e)}, 401)

    try:
        body = req.get_json()
    except ValueError:
        return _json({"error": "Expected a JSON body"}, 400)

    chosen = body.get("chosen", {})
    if not chosen.get("driveId") or not chosen.get("itemId"):
        return _json({"error": "chosen.driveId and chosen.itemId are required"}, 400)
    # A site-level pseudo target has no real drive to upload into.
    if str(chosen.get("driveId", "")).startswith("site:") or chosen.get("itemId") == "drives":
        return _json({"error": "Pick a folder inside a document library, not a site."}, 400)

    mode = body.get("mode", "single")

    # Accept BOTH payload shapes: the new `items` array, and the legacy `document` object.
    items = body.get("items")
    if not items:
        doc = body.get("document") or {}
        items = [doc] if doc else []

    files = [i for i in items if i.get("kind", "file") == "file"]
    emails = [i for i in items if i.get("kind") == "email"]

    if not files and not emails:
        return _json({"error": "Nothing to file: no attachments or email in the request."}, 400)

    uploaded, skipped = [], []
    for it in files:
        name = it.get("name") or "filed-document"
        content_b64 = it.get("contentBytes")
        if not content_b64:
            # The bytes never arrived from Office.js -- do NOT pretend this worked.
            logging.warning("recordChoice: no contentBytes for %r -- skipping upload", name)
            skipped.append(name)
            continue
        try:
            content = base64.b64decode(content_b64)
            graph_client.upload_file(
                graph_token, chosen["driveId"], chosen["itemId"], name, content,
            )
            uploaded.append(name)
            logging.info("recordChoice: uploaded %r (%d bytes)", name, len(content))
        except (graph_client.GraphError, ValueError) as e:
            logging.warning("recordChoice upload %r: %s", name, e)
            return _json({"error": f"Upload failed for {name}: {e}"}, 502)

    # Saving the email itself (.pst) is not implemented server-side yet.
    if emails and not files:
        return _json(
            {"error": "Saving the email as .pst is not implemented on the backend yet."}, 501
        )

    # If we were asked to file documents but uploaded none, that is a failure -- say so.
    if files and not uploaded:
        return _json(
            {
                "error": "Could not read the attachment contents, so nothing was uploaded. "
                         "Reopen the task pane and try again.",
                "skipped": skipped,
            },
            400,
        )

    # Log one decision row per uploaded document (the training signal).
    for it in files:
        name = it.get("name") or "filed-document"
        if name in skipped:
            continue
        storage.append_decision({
            "documentName": name,
            "subject": it.get("subject") or body.get("subject"),
            "messageId": it.get("messageId"),
            "chosenPath": chosen.get("path"),
            "chosenDriveId": chosen.get("driveId"),
            "chosenItemId": chosen.get("itemId"),
            "site": chosen.get("site"),
            "recommendations": body.get("recommendations", []),
            "acceptedRecommendation": bool(body.get("acceptedRecommendation")),
            "acceptedRank": body.get("acceptedRank"),
            "timestamp": body.get("timestamp"),
            "mode": mode,
        })

    return _json({"status": "ok", "uploaded": uploaded, "skipped": skipped})


