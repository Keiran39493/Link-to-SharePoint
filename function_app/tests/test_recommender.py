import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from recommender import recommend  # noqa: E402


def _row(name, subject, path, drive_id="d1", item_id="i1", site="Corporate"):
    return {
        "documentName": name,
        "subject": subject,
        "chosenPath": path,
        "chosenDriveId": drive_id,
        "chosenItemId": item_id,
        "site": site,
    }


def test_no_history_returns_empty():
    assert recommend({"name": "HMRC_letter_2026.pdf", "subject": "Q3 tax"}, []) == []


def test_no_textual_overlap_returns_empty():
    history = [_row("staff_handbook.docx", "HR policy update", "SharePoint/HR/Policies")]
    doc = {"name": "invoice_9471.pdf", "subject": "Overdue payment reminder"}
    assert recommend(doc, history) == []


def test_matches_similar_document_to_its_folder():
    history = [
        _row("HMRC_letter_2025.pdf", "Q2 tax correspondence Galvan Ltd",
             "SharePoint/Corporate/TAX/HMRC_correspondence/Galvan_Ltd", item_id="galvan"),
        _row("Board_minutes_2025.pdf", "Board meeting minutes", "SharePoint/Corporate/OVERVIEW", item_id="ov"),
    ]
    doc = {"name": "HMRC_letter_2026.pdf", "subject": "Q3 tax correspondence Galvan Ltd"}

    recs = recommend(doc, history)

    assert recs, "expected at least one recommendation"
    assert recs[0]["path"] == "SharePoint/Corporate/TAX/HMRC_correspondence/Galvan_Ltd"
    assert recs[0]["itemId"] == "galvan"
    assert 0 < recs[0]["confidence"] <= 0.97
    assert "1 similar document" in recs[0]["reason"]


def test_ranks_by_relevance_descending():
    history = [
        _row("HMRC_letter_2024.pdf", "tax correspondence Galvan Ltd", "SharePoint/TAX/Galvan", item_id="galvan"),
        _row("random_memo.docx", "office relocation notice", "SharePoint/Facilities", item_id="fac"),
    ]
    doc = {"name": "HMRC_letter_2026.pdf", "subject": "tax correspondence Galvan Ltd"}

    recs = recommend(doc, history)

    confidences = [r["confidence"] for r in recs]
    assert confidences == sorted(confidences, reverse=True)
    assert recs[0]["itemId"] == "galvan"


def test_repetition_increases_confidence_over_single_match():
    doc = {"name": "HMRC_letter_2026.pdf", "subject": "tax correspondence Galvan Ltd"}

    single = recommend(doc, [
        _row("HMRC_letter_2025.pdf", "tax correspondence Galvan Ltd", "A", item_id="a"),
    ])
    repeated = recommend(doc, [
        _row("HMRC_letter_2025.pdf", "tax correspondence Galvan Ltd", "B", item_id="b"),
        _row("HMRC_letter_2024.pdf", "tax correspondence Galvan Ltd", "B", item_id="b"),
        _row("HMRC_letter_2023.pdf", "tax correspondence Galvan Ltd", "B", item_id="b"),
    ])

    assert repeated[0]["confidence"] > single[0]["confidence"]


def test_confidence_never_exceeds_cap():
    doc = {"name": "HMRC_letter_2026.pdf", "subject": "tax correspondence Galvan Ltd"}
    history = [_row("HMRC_letter_2026.pdf", "tax correspondence Galvan Ltd", "A", item_id="a")] * 50

    recs = recommend(doc, history)

    assert recs[0]["confidence"] <= 0.97


def test_respects_top_n():
    doc = {"name": "letter.pdf", "subject": "quarterly report"}
    history = [
        _row("letter_a.pdf", "quarterly report alpha", "A", item_id="a"),
        _row("letter_b.pdf", "quarterly report beta", "B", item_id="b"),
        _row("letter_c.pdf", "quarterly report gamma", "C", item_id="c"),
        _row("letter_d.pdf", "quarterly report delta", "D", item_id="d"),
    ]

    assert len(recommend(doc, history, top_n=2)) == 2
    assert len(recommend(doc, history)) <= 3
