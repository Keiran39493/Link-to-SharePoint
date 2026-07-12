"""The training log: one row per recordChoice call.

Reuses the Function App's own AzureWebJobsStorage account (every Function
App already has one) so there's no second connection string to provision.

All rows share one PartitionKey ("decision"). Write volume here is "one
per human filing an email" -- nowhere near Table Storage's per-partition
throughput ceiling -- and a single partition is what makes "give me the
most recent N decisions" a cheap, correctly-ordered query: RowKey is an
inverted timestamp, so ascending RowKey order (Table Storage's native
order within a partition) is reverse-chronological. Splitting by site/date
would break that global ordering without buying anything at this scale.
"""
from __future__ import annotations

import datetime
import itertools
import json
import os
import uuid
from typing import Any

from azure.data.tables import TableServiceClient

CONNECTION_STRING = os.environ["AzureWebJobsStorage"]
TABLE_NAME = os.environ.get("FILING_TABLE_NAME", "FilingDecisions")

_PARTITION_KEY = "decision"
_JSON_FIELDS = {"recommendations"}
_INVERTED_TS_CEILING = 9_999_999_999_999  # ms epoch; good until year ~2286


def _table_client():
    service = TableServiceClient.from_connection_string(CONNECTION_STRING)
    service.create_table_if_not_exists(TABLE_NAME)
    return service.get_table_client(TABLE_NAME)


def _row_key() -> str:
    now_ms = int(datetime.datetime.now(datetime.timezone.utc).timestamp() * 1000)
    inverted = _INVERTED_TS_CEILING - now_ms
    return f"{inverted:013d}-{uuid.uuid4().hex[:8]}"


def append_decision(row: dict[str, Any]) -> None:
    """Logs one filing decision. `row` should have: documentName, subject,
    messageId, chosenPath, chosenDriveId, chosenItemId, site,
    recommendations, acceptedRecommendation, acceptedRank, timestamp."""
    entity: dict[str, Any] = {"PartitionKey": _PARTITION_KEY, "RowKey": _row_key()}
    for key, value in row.items():
        if key in _JSON_FIELDS or not isinstance(value, (str, int, float, bool, type(None))):
            value = json.dumps(value)
        entity[key] = value
    _table_client().create_entity(entity)


def fetch_recent_decisions(limit: int = 500) -> list[dict[str, Any]]:
    """Most recent decisions first, bounded so recommend() stays fast as
    the log grows. Relies on the inverted-timestamp RowKey (see module
    docstring) -- only pages in as many rows as `limit` needs."""
    client = _table_client()
    rows = itertools.islice(client.list_entities(results_per_page=limit), limit)
    decoded = []
    for row in rows:
        item = dict(row)
        for key in _JSON_FIELDS:
            if key in item and isinstance(item[key], str):
                try:
                    item[key] = json.loads(item[key])
                except (json.JSONDecodeError, TypeError):
                    pass
        decoded.append(item)
    return decoded
