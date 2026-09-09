"""Closed value contracts for Windows Microsoft session business calls.

Only explicitly listed public WPSComposer value types cross the JSON worker
boundary.  Native objects use the separate session-handle registry.
"""
from __future__ import annotations

import datetime as _datetime
import math
from dataclasses import fields

from ..design_presets import DesignPreset
from ..document_model import (
    CitationRun,
    CrossReferenceRun,
    InlineDegradationRun,
    Span,
)
from ..layout_templates import LayoutTemplate
from ..longform.executor import ExecutionIssue, FieldSnapshot
from ..longform.quality import QualityFinding


DTO_TAG = "__wpscomposer_dto__"


_DATACLASS_TYPES = {
    cls.__name__: cls
    for cls in (
        CitationRun,
        CrossReferenceRun,
        InlineDegradationRun,
        Span,
    )
}
_MAPPING_TYPES = {
    "ExecutionIssue": (
        ExecutionIssue,
        ExecutionIssue.from_dict,
        frozenset({"code", "message"}),
        frozenset({"placement", "nodeId", "stage", "fallback", "recoverable"}),
    ),
    "FieldSnapshot": (
        FieldSnapshot,
        FieldSnapshot.from_dict,
        frozenset({
            "stableKey", "fieldCategory", "resultHash", "tocPageCount",
            "figureIndexPageCount", "tableIndexPageCount", "totalPages",
        }),
        frozenset(),
    ),
    "QualityFinding": (
        QualityFinding,
        QualityFinding.from_dict,
        frozenset({"code", "severity", "confidence", "message"}),
        frozenset({"nodeId", "page", "bounds", "evidence", "repairKey"}),
    ),
}


def _tag(name, value):
    return {DTO_TAG: {"type": name, "value": value}}


def encode_value(value):
    """Encode one closed public value into JSON-only data."""
    if value is None or type(value) in (str, int, bool):
        return value
    if type(value) is float:
        if not math.isfinite(value):
            raise ValueError("Business protocol numbers must be finite")
        return value
    # pywintypes.TimeType can be a datetime subclass.  Preserve the public
    # date/datetime category without importing or naming the extension class.
    if isinstance(value, _datetime.datetime):
        return _tag("datetime", value.isoformat())
    if isinstance(value, _datetime.date):
        return _tag("date", value.isoformat())
    if type(value) is tuple:
        return _tag("tuple", [encode_value(item) for item in value])
    if type(value) is list:
        return [encode_value(item) for item in value]
    if type(value) is dict:
        return {key: encode_value(item) for key, item in value.items()}
    if type(value) is DesignPreset:
        return _tag(
            "DesignPreset",
            encode_value({
                "name": value.name,
                "colors": value.colors,
                "fonts": value.fonts,
                "spacing": value.spacing,
                "rules": value.rules,
            }),
        )
    if type(value) is LayoutTemplate:
        return _tag(
            "LayoutTemplate",
            encode_value({
                "name": value.name,
                "category": value.category,
                "description": value.description,
                "elements": value.elements,
                "structural_rules": value.structural_rules,
            }),
        )
    for name, cls in _DATACLASS_TYPES.items():
        if type(value) is cls:
            return _tag(
                name,
                {field.name: encode_value(getattr(value, field.name)) for field in fields(cls)},
            )
    for name, (cls, _decoder, _required, _optional) in _MAPPING_TYPES.items():
        if type(value) is cls:
            return _tag(name, encode_value(value.to_dict()))
    raise TypeError(
        f"Object of type {type(value).__name__} is outside the closed business protocol"
    )


def _mapping(value, expected):
    if not isinstance(value, dict) or set(value) != set(expected):
        raise ValueError("Invalid business DTO fields")
    return value


def _mapping_fields(value, required, optional):
    if not isinstance(value, dict):
        raise ValueError("Invalid business DTO fields")
    present = set(value)
    if not required <= present or not present <= required | optional:
        raise ValueError("Invalid business DTO fields")
    return value


def decode_value(value):
    """Reconstruct one value previously produced by :func:`encode_value`."""
    if value is None or type(value) in (str, int, bool):
        return value
    if type(value) is float:
        if not math.isfinite(value):
            raise ValueError("Business protocol numbers must be finite")
        return value
    if type(value) is list:
        return [decode_value(item) for item in value]
    if type(value) is not dict:
        raise ValueError("Invalid business protocol value")
    if set(value) != {DTO_TAG}:
        return {key: decode_value(item) for key, item in value.items()}
    envelope = value[DTO_TAG]
    if not isinstance(envelope, dict) or set(envelope) != {"type", "value"}:
        raise ValueError("Invalid business DTO envelope")
    name = envelope["type"]
    payload = envelope["value"]
    if name == "tuple":
        if type(payload) is not list:
            raise ValueError("Invalid business DTO tuple")
        return tuple(decode_value(item) for item in payload)
    if name == "date":
        if type(payload) is not str:
            raise ValueError("Invalid business DTO date")
        try:
            return _datetime.date.fromisoformat(payload)
        except ValueError:
            raise ValueError("Invalid business DTO date") from None
    if name == "datetime":
        if type(payload) is not str:
            raise ValueError("Invalid business DTO datetime")
        try:
            return _datetime.datetime.fromisoformat(payload)
        except ValueError:
            raise ValueError("Invalid business DTO datetime") from None
    decoded = decode_value(payload)
    if name == "DesignPreset":
        data = _mapping(decoded, ("name", "colors", "fonts", "spacing", "rules"))
        return DesignPreset(**data)
    if name == "LayoutTemplate":
        data = _mapping(
            decoded,
            ("name", "category", "description", "elements", "structural_rules"),
        )
        return LayoutTemplate(**data)
    if name in _DATACLASS_TYPES:
        cls = _DATACLASS_TYPES[name]
        data = _mapping(decoded, tuple(field.name for field in fields(cls)))
        try:
            return cls(**data)
        except (TypeError, ValueError, KeyError):
            raise ValueError("Invalid business DTO fields") from None
    if name in _MAPPING_TYPES:
        _cls, decoder, required, optional = _MAPPING_TYPES[name]
        decoded = _mapping_fields(decoded, required, optional)
        try:
            return decoder(decoded)
        except (TypeError, ValueError, KeyError, IndexError):
            raise ValueError("Invalid business DTO fields") from None
    raise ValueError("Unknown business DTO type")


__all__ = ["DTO_TAG", "decode_value", "encode_value"]
