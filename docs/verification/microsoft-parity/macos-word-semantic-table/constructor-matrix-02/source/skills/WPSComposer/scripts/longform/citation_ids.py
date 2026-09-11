"""Deterministic identities for explicit table-cell citation relations."""

from __future__ import annotations

import hashlib
import json


def table_cell_citation_node_id(
    table_node_id: str, row: int, column: int, target_id: str
) -> str:
    payload = json.dumps(
        [table_node_id, row, column, target_id],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]
    return f"{table_node_id}/cell:{row}:{column}/cite:{digest}"


__all__ = ["table_cell_citation_node_id"]
