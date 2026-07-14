"""Lightweight, self-hosted recommender for filing destinations.

No AI Builder / Copilot Studio license required: this scores candidate
folders (drawn from the historical filing log) by text overlap between the
new document's name/subject and past documents filed to each folder, plus
a small bonus for how often that folder gets chosen for similar mail. Every
recordChoice call adds one more data point, so it genuinely learns.

Swap this module out later for real AI Builder / Copilot Studio ranking if
you want it -- recommend() is the only thing function_app.py depends on.
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

_STOPWORDS = {
    "the", "a", "an", "and", "or", "of", "for", "to", "in", "on", "re",
    "fw", "fwd", "pdf", "doc", "docx", "letter", "email", "attachment",
}
_TOKEN_RE = re.compile(r"[a-z0-9]+")
_MIN_CONFIDENCE = 0.15


def _tokens(*texts: str | None) -> set[str]:
    joined = " ".join(t for t in texts if t)
    return {
        tok for tok in _TOKEN_RE.findall(joined.lower())
        if len(tok) > 2 and tok not in _STOPWORDS
    }


def _similarity(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    overlap = len(a & b)
    if not overlap:
        return 0.0
    return overlap / len(a | b)  # Jaccard


def recommend(document: dict[str, Any], history: list[dict[str, Any]], top_n: int = 3) -> list[dict[str, Any]]:
    """Rank past filing destinations for a new document.

    `document` needs `name`/`subject`. `history` is a list of past
    recordChoice rows (see storage.append_decision) -- each needs
    chosenPath/chosenDriveId/chosenItemId/site/documentName/subject.
    Returns up to `top_n` entries shaped like the /sp/recommend response.
    """
    new_tokens = _tokens(document.get("name"), document.get("subject"))
    if not new_tokens or not history:
        return []

    groups: dict[tuple[Any, Any, Any], dict[str, Any]] = defaultdict(
        lambda: {"count": 0, "best_sim": 0.0, "sim_sum": 0.0, "site": None, "webUrl": None}
    )
    for row in history:
        path = row.get("chosenPath")
        drive_id = row.get("chosenDriveId")
        item_id = row.get("chosenItemId")
        if not path or drive_id is None or item_id is None:
            continue
        key = (path, drive_id, item_id)
        sim = _similarity(new_tokens, _tokens(row.get("documentName"), row.get("subject")))
        g = groups[key]
        g["count"] += 1
        g["sim_sum"] += sim
        g["best_sim"] = max(g["best_sim"], sim)
        g["site"] = row.get("site")
        g["webUrl"] = row.get("chosenWebUrl")

    scored = []
    for (path, drive_id, item_id), g in groups.items():
        if g["best_sim"] <= 0:
            continue  # no textual signal at all -- don't guess
        avg_sim = g["sim_sum"] / g["count"]
        # Blend the single best match with the average and a small bump for
        # repetition, so one strong hit plus a history of reuse beats a
        # single fluke match. Capped short of 1.0 -- never claim certainty.
        confidence = min(0.97, 0.6 * g["best_sim"] + 0.3 * avg_sim + min(0.1, 0.02 * g["count"]))
        if confidence < _MIN_CONFIDENCE:
            continue
        scored.append({
            "path": path,
            "site": g["site"],
            "driveId": drive_id,
            "itemId": item_id,
            "confidence": round(confidence, 2),
            "reason": f"{g['count']} similar document{'s' if g['count'] != 1 else ''} filed here previously.",
        })

    scored.sort(key=lambda r: r["confidence"], reverse=True)
    return scored[:top_n]
