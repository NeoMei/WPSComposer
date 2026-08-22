"""Page-skeleton policy for long-form documents.

Determines the ordered page-role sequence, page-numbering rules,
header/footer flags, and link-to-previous rules from the resolved semantic
model and policy.  This module never touches platform objects.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from ..document_model import StructuredDocument
from .policy import LongformPolicy
from .semantic import LongformConfig


@dataclass(frozen=True)
class PageSectionPolicy:
    """Policy for one page-skeleton section/role."""

    role: str
    page_number_format: str
    start_page_number: Optional[int]
    restart_numbering: bool
    has_header: bool
    has_footer: bool
    header_text: str
    footer_text: str
    link_to_previous_header: bool
    link_to_previous_footer: bool
    includes_abstract: bool = False
    includes_keywords: bool = False
    includes_toc: bool = False
    includes_figure_index: bool = False
    includes_table_index: bool = False


@dataclass(frozen=True)
class PageSkeletonPolicy:
    """Ordered page-role policy for a document."""

    sections: tuple[PageSectionPolicy, ...]


def _role_policy(
    policy: LongformPolicy,
    role: str,
    **overrides: Any,
) -> PageSectionPolicy:
    """Build a PageSectionPolicy from the policy numbering rules and overrides."""
    rules = policy.section_numbering_rules.get(role, {})
    defaults = {
        "role": role,
        "page_number_format": rules.get("format", "none"),
        "start_page_number": rules.get("start_page_number"),
        "restart_numbering": bool(rules.get("restart", False)),
        "has_header": False,
        "has_footer": False,
        "header_text": "",
        "footer_text": "",
        "link_to_previous_header": False,
        "link_to_previous_footer": False,
    }
    defaults.update(overrides)
    return PageSectionPolicy(**defaults)


def _has_front_matter(doc: StructuredDocument, policy: LongformPolicy) -> bool:
    """Return True if any front-matter content is enabled/present."""
    return bool(
        doc.abstract
        or doc.keywords
        or policy.toc
        or policy.figure_index
        or policy.table_index
    )


def _section_role(section: Any) -> str:
    """Return the page role for a document section."""
    return "landscape" if getattr(section, "orientation", None) == "landscape" else "body"


def _content_role(role: str) -> bool:
    """Return True if *role* is a content role that may carry headers/footers."""
    return role in {"body", "landscape"}


def _group_sections(sections: list) -> list[tuple[str, list]]:
    """Group consecutive document sections by orientation role in order."""
    groups: list[tuple[str, list]] = []
    current_role: Optional[str] = None
    current: list = []
    for section in sections or []:
        role = _section_role(section)
        if role != current_role:
            if current_role is not None:
                groups.append((current_role, current))
            current_role = role
            current = [section]
        else:
            current.append(section)
    if current_role is not None:
        groups.append((current_role, current))
    return groups


def build_page_policy(
    doc: StructuredDocument,
    config: LongformConfig,
    policy: LongformPolicy,
) -> PageSkeletonPolicy:
    """Return the ordered page-skeleton policy for *doc*.

    The result is deterministic and never raises; malformed input produces an
    empty skeleton so that callers can continue degrading safely.
    """
    try:
        sections: list[PageSectionPolicy] = []

        # Cover page is present only when explicitly enabled and a title exists.
        if policy.title_page and policy.title:
            sections.append(
                _role_policy(
                    policy,
                    "cover",
                    has_header=False,
                    has_footer=False,
                )
            )

        # Front matter groups abstract, keywords, TOC, and optional indexes.
        if _has_front_matter(doc, policy):
            sections.append(
                _role_policy(
                    policy,
                    "front_matter",
                    has_footer=True,
                    includes_abstract=bool(doc.abstract),
                    includes_keywords=bool(doc.keywords),
                    includes_toc=policy.toc,
                    includes_figure_index=policy.figure_index,
                    includes_table_index=policy.table_index,
                )
            )

        # Content sections preserve document order and group only consecutive
        # same-orientation sections.
        for role, _ in _group_sections(doc.sections):
            previous_role = sections[-1].role if sections else None
            link = _content_role(previous_role)
            if role == "body":
                sections.append(
                    _role_policy(
                        policy,
                        "body",
                        has_header=True,
                        has_footer=True,
                        header_text=policy.header_text,
                        link_to_previous_header=link,
                        link_to_previous_footer=link,
                    )
                )
            elif role == "landscape":
                sections.append(
                    _role_policy(
                        policy,
                        "landscape",
                        has_header=True,
                        has_footer=True,
                        header_text=policy.header_text,
                        link_to_previous_header=link,
                        link_to_previous_footer=link,
                    )
                )

        return PageSkeletonPolicy(tuple(sections))
    except Exception:
        # Deterministic degradation: never raise on bad input.
        return PageSkeletonPolicy(())


__all__ = [
    "PageSectionPolicy",
    "PageSkeletonPolicy",
    "build_page_policy",
]
