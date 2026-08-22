"""Focused schema tests for M2 protocol v2 page-skeleton operations."""

from __future__ import annotations

import pytest
from skills.WPSComposer.scripts.generation_plan import (
    OperationPlanError,
    validate_generation_plan,
)


def _v2_plan(*operations):
    return {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": "sha256:" + "0" * 64,
        "operations": list(operations),
    }


def _set_page_role(role: str, node_id: str = "sec:body"):
    return {
        "op": "writer.set_page_role",
        "nodeId": node_id,
        "args": {"role": role},
    }


def _set_page_numbering(
    format: str,
    start: int | None = None,
    restart: bool = False,
    node_id: str = "sec:body",
):
    args: dict = {"format": format}
    if start is not None:
        args["start"] = start
    if restart:
        args["restart"] = restart
    return {
        "op": "writer.set_page_numbering",
        "nodeId": node_id,
        "args": args,
    }


def _set_header_footer(
    header_text: str = "",
    footer_text: str = "",
    link_header: bool = False,
    link_footer: bool = False,
    node_id: str = "sec:body",
):
    return {
        "op": "writer.set_header_footer",
        "nodeId": node_id,
        "args": {
            "headerText": header_text,
            "footerText": footer_text,
            "linkToPreviousHeader": link_header,
            "linkToPreviousFooter": link_footer,
        },
    }


class TestPageSkeletonOperations:
    """Schema validation for writer.set_page_role / numbering / header_footer."""

    def test_set_page_role_requires_node_id(self):
        raw = _v2_plan({"op": "writer.set_page_role", "args": {"role": "body"}})
        with pytest.raises(OperationPlanError, match="nodeId"):
            validate_generation_plan(raw, "writer")

    def test_set_page_numbering_requires_node_id(self):
        raw = _v2_plan(
            {"op": "writer.set_page_numbering", "args": {"format": "arabic"}}
        )
        with pytest.raises(OperationPlanError, match="nodeId"):
            validate_generation_plan(raw, "writer")

    def test_set_header_footer_requires_node_id(self):
        raw = _v2_plan(
            {"op": "writer.set_header_footer", "args": {"headerText": "H"}}
        )
        with pytest.raises(OperationPlanError, match="nodeId"):
            validate_generation_plan(raw, "writer")

    def test_set_page_role_rejects_unknown_field(self):
        raw = _v2_plan(
            {
                "op": "writer.set_page_role",
                "nodeId": "sec:body",
                "args": {"role": "body", "unknownField": 1},
            }
        )
        with pytest.raises(OperationPlanError, match="unknown argument"):
            validate_generation_plan(raw, "writer")

    def test_set_page_role_rejects_invalid_role(self):
        raw = _v2_plan(_set_page_role("not-a-role"))
        with pytest.raises(OperationPlanError, match="role"):
            validate_generation_plan(raw, "writer")

    def test_set_page_numbering_rejects_unknown_field(self):
        raw = _v2_plan(
            {
                "op": "writer.set_page_numbering",
                "nodeId": "sec:body",
                "args": {"format": "arabic", "extra": True},
            }
        )
        with pytest.raises(OperationPlanError, match="unknown argument"):
            validate_generation_plan(raw, "writer")

    @pytest.mark.parametrize("fmt", ["roman", "arabic", "none", "continue"])
    def test_set_page_numbering_accepts_valid_formats(self, fmt):
        raw = _v2_plan(_set_page_numbering(fmt, start=1, restart=True))
        plan = validate_generation_plan(raw, "writer")
        op = plan.operations[0]
        assert op.op == "writer.set_page_numbering"
        assert op.args["format"] == fmt

    @pytest.mark.parametrize("fmt", ["alphabetic", "Roman", "123", "arabic-roman"])
    def test_set_page_numbering_rejects_invalid_format(self, fmt):
        raw = _v2_plan(_set_page_numbering(fmt))
        with pytest.raises(OperationPlanError, match="format"):
            validate_generation_plan(raw, "writer")

    def test_set_header_footer_rejects_unknown_field(self):
        raw = _v2_plan(
            {
                "op": "writer.set_header_footer",
                "nodeId": "sec:body",
                "args": {"headerText": "H", "footerText": "F", "bad": 1},
            }
        )
        with pytest.raises(OperationPlanError, match="unknown argument"):
            validate_generation_plan(raw, "writer")


class TestConfigureSectionSchemaExtension:
    """Schema validation for the M2 writer.configure_section extension."""

    def test_configure_section_accepts_page_numbering_and_header_footer_fields(self):
        raw = _v2_plan(
            {
                "op": "writer.configure_section",
                "nodeId": "sec:body",
                "args": {
                    "landscape": False,
                    "pageNumberFormat": "arabic",
                    "restartPageNumbering": True,
                    "startPageNumber": 1,
                    "headerText": "Header",
                    "footerText": "Footer",
                    "linkToPreviousHeader": False,
                    "linkToPreviousFooter": False,
                },
            }
        )
        plan = validate_generation_plan(raw, "writer")
        args = plan.operations[0].args
        assert args["pageNumberFormat"] == "arabic"
        assert args["restartPageNumbering"] is True
        assert args["linkToPreviousFooter"] is False

    def test_configure_section_rejects_invalid_page_number_format(self):
        raw = _v2_plan(
            {
                "op": "writer.configure_section",
                "nodeId": "sec:body",
                "args": {"pageNumberFormat": "chinese-formal"},
            }
        )
        with pytest.raises(OperationPlanError, match="pageNumberFormat"):
            validate_generation_plan(raw, "writer")

    def test_configure_section_rejects_unknown_field(self):
        raw = _v2_plan(
            {
                "op": "writer.configure_section",
                "nodeId": "sec:body",
                "args": {"pageNumberFormat": "arabic", "extra": "bad"},
            }
        )
        with pytest.raises(OperationPlanError, match="unknown argument"):
            validate_generation_plan(raw, "writer")


class TestTocDensitySchema:
    """Schema validation for writer.configure_toc_styles density fields."""

    def _density_args(
        self,
        min_font_size_toc1: float = 10.5,
        min_font_size_toc2: float = 10.0,
        min_font_size_toc3: float = 10.0,
        min_space_before: float = 0.0,
        min_space_after: float = 0.0,
    ) -> dict:
        return {
            "tocTitle": "目录",
            "levels": 3,
            "minFontSizePt": {
                "toc1": min_font_size_toc1,
                "toc2": min_font_size_toc2,
                "toc3": min_font_size_toc3,
            },
            "minSpaceBeforePt": {
                "toc1": min_space_before,
                "toc2": min_space_before,
                "toc3": min_space_before,
            },
            "minSpaceAfterPt": {
                "toc1": min_space_after,
                "toc2": min_space_after,
                "toc3": min_space_after,
            },
        }

    def test_configure_toc_styles_accepts_minimum_density(self):
        raw = _v2_plan(
            {
                "op": "writer.configure_toc_styles",
                "nodeId": "doc:toc",
                "args": self._density_args(),
            }
        )
        plan = validate_generation_plan(raw, "writer")
        args = plan.operations[0].args
        assert args["minFontSizePt"]["toc1"] == 10.5

    @pytest.mark.parametrize(
        ("field", "level", "value", "bound"),
        [
            ("minFontSizePt", "toc1", 10.4, 10.5),
            ("minFontSizePt", "toc2", 9.9, 10.0),
            ("minFontSizePt", "toc3", 9.9, 10.0),
            ("minSpaceBeforePt", "toc1", -0.1, 0.0),
            ("minSpaceAfterPt", "toc2", -0.1, 0.0),
        ],
    )
    def test_configure_toc_styles_rejects_density_below_lower_bound(
        self, field, level, value, bound
    ):
        args = self._density_args()
        args[field][level] = value
        raw = _v2_plan({"op": "writer.configure_toc_styles", "nodeId": "doc:toc", "args": args})
        with pytest.raises(OperationPlanError, match=f"{field}\\.{level}|{field}"):
            validate_generation_plan(raw, "writer")

    def test_configure_toc_styles_rejects_unknown_density_field(self):
        args = self._density_args()
        args["minFontSizePt"]["toc4"] = 10.0
        raw = _v2_plan({"op": "writer.configure_toc_styles", "nodeId": "doc:toc", "args": args})
        with pytest.raises(OperationPlanError, match="minFontSizePt"):
            validate_generation_plan(raw, "writer")


class TestHeadingNumberingSchema:
    """Schema validation for writer.add_heading numbering fields."""

    def test_add_heading_accepts_numbering_and_scheme(self):
        raw = {
            "component": "writer",
            "operations": [
                {
                    "op": "writer.add_heading",
                    "args": {
                        "text": "第一章",
                        "level": 1,
                        "numbering": True,
                        "numberingScheme": "chinese-formal",
                    },
                }
            ],
        }
        plan = validate_generation_plan(raw, "writer")
        args = plan.operations[0].args
        assert args["numbering"] is True
        assert args["numberingScheme"] == "chinese-formal"

    def test_add_heading_rejects_unknown_numbering_scheme(self):
        raw = {
            "component": "writer",
            "operations": [
                {
                    "op": "writer.add_heading",
                    "args": {
                        "text": "第一章",
                        "level": 1,
                        "numbering": True,
                        "numberingScheme": "alpha",
                    },
                }
            ],
        }
        with pytest.raises(OperationPlanError, match="numberingScheme"):
            validate_generation_plan(raw, "writer")

    def test_add_heading_rejects_unknown_numbering_field(self):
        raw = {
            "component": "writer",
            "operations": [
                {
                    "op": "writer.add_heading",
                    "args": {
                        "text": "第一章",
                        "level": 1,
                        "numbering": True,
                        "numberingScheme": "decimal",
                        "extra": 1,
                    },
                }
            ],
        }
        with pytest.raises(OperationPlanError, match="unknown argument"):
            validate_generation_plan(raw, "writer")
