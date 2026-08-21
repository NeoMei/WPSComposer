from __future__ import annotations

from dataclasses import dataclass
from typing import Any, List, Optional, Tuple

from ..document_model import (
    StructuredDocument, Section, Paragraph, Span, ListBlock, TableBlock,
    ImageBlock, ExcalidrawBlock, DocumentIssue, AbstractBlock, KeywordsBlock,
    PageBreakBlock, SemanticTableBlock, FigureBlock, FormulaBlock,
    ReferenceListBlock, DegradationBlock,
)
from .frontmatter_parser import parse_frontmatter_document
from .directives import (
    scan_block_directives, BlockDirective,
    DIRECTIVE_SYNTAX_INVALID, DIRECTIVE_UNCLOSED, NESTED_DIRECTIVE_UNSUPPORTED,
)


LONGFORM_DIRECTIVE_UNKNOWN = "LONGFORM_DIRECTIVE_UNKNOWN"
LONGFORM_DUPLICATE_FRONT_BLOCK = "LONGFORM_DUPLICATE_FRONT_BLOCK"
LONGFORM_FRONTMATTER_VALUE_IGNORED = "LONGFORM_FRONTMATTER_VALUE_IGNORED"


@dataclass
class DirectiveContext:
    doc: StructuredDocument
    seen_front_blocks: set[str]
    sections: List[Section]
    current_section: Optional[Section]
    base_dir: str


def _ensure_section(current: Optional[Section], sections: List[Section]) -> Section:
    if current is None:
        current = Section(level=0, heading="")
        sections.append(current)
    return current


def _paragraphs_from_text(parse_inline, text: str) -> List[Paragraph]:
    paragraphs: List[Paragraph] = []
    for raw in text.split("\n\n"):
        raw = raw.strip()
        if raw:
            paragraphs.append(Paragraph(spans=parse_inline(raw)))
    return paragraphs


def _abstract_from_value(parse_inline, value: Any) -> Optional[AbstractBlock]:
    if isinstance(value, str):
        return AbstractBlock(paragraphs=_paragraphs_from_text(parse_inline, value))
    if isinstance(value, list):
        paragraphs: List[Paragraph] = []
        for item in value:
            if isinstance(item, str):
                paragraphs.extend(_paragraphs_from_text(parse_inline, item))
        return AbstractBlock(paragraphs=paragraphs)
    return None


def _keywords_from_value(value: Any) -> Optional[KeywordsBlock]:
    if isinstance(value, str):
        parts = [p.strip() for p in value.replace(",", "\n").split("\n") if p.strip()]
        return KeywordsBlock(keywords=parts)
    if isinstance(value, list):
        parts = [str(item).strip() for item in value if str(item).strip()]
        return KeywordsBlock(keywords=parts)
    return None


def _frontmatter_value_to_str(value: Any) -> Tuple[str, bool]:
    if isinstance(value, str):
        return value, True
    if value is None:
        return "", True
    if isinstance(value, (bool, int, float)):
        return str(value), True
    return str(value), False


def _collect_elements(parse_block_lines, lines: List[str], base_dir: str = "") -> List[Any]:
    sections, _ = parse_block_lines(lines, base_dir=base_dir)
    elements: List[Any] = []
    for sec in sections:
        elements.extend(sec.elements)
    return elements


def parse_longform(
    md_text: str,
    base_dir: str,
    parse_inline,
    parse_block_lines,
    detect_first_h1,
) -> StructuredDocument:
    doc = StructuredDocument(longform=True)
    seen_front_blocks: set[str] = set()

    fm_result = parse_frontmatter_document(md_text)
    for issue in fm_result.issues:
        doc.issues.append(DocumentIssue(code=issue, message=issue, placement="document"))
    body = fm_result.body

    if "title" in fm_result.values and isinstance(fm_result.values["title"], str):
        doc.title = fm_result.values["title"]

    if "abstract" in fm_result.values:
        abstract = _abstract_from_value(parse_inline, fm_result.values["abstract"])
        if abstract is not None:
            doc.abstract = abstract
        else:
            doc.issues.append(DocumentIssue(
                code=LONGFORM_FRONTMATTER_VALUE_IGNORED,
                message="Frontmatter 'abstract' value ignored: must be a string or list of strings.",
                placement="document",
            ))

    if "keywords" in fm_result.values:
        keywords = _keywords_from_value(fm_result.values["keywords"])
        if keywords is not None:
            doc.keywords = keywords
        else:
            doc.issues.append(DocumentIssue(
                code=LONGFORM_FRONTMATTER_VALUE_IGNORED,
                message="Frontmatter 'keywords' value ignored: must be a string or list of strings.",
                placement="document",
            ))

    for key, value in fm_result.values.items():
        if key in ("title", "abstract", "keywords"):
            continue
        text_value, scalar = _frontmatter_value_to_str(value)
        if scalar:
            doc.metadata[key] = text_value
        else:
            doc.issues.append(DocumentIssue(
                code=LONGFORM_FRONTMATTER_VALUE_IGNORED,
                message=f"Frontmatter '{key}' is not a scalar value and was ignored.",
                placement="document",
            ))

    regions = scan_block_directives(body)
    sections: List[Section] = []
    current_section: Optional[Section] = None

    for region in regions:
        if isinstance(region, str):
            if region:
                lines = region.split("\n")
                sections, current_section = parse_block_lines(
                    lines, base_dir=base_dir, sections=sections, current_section=current_section
                )
            continue

        if isinstance(region, BlockDirective):
            directive = region
            if directive.issues:
                for code in directive.issues:
                    current_section = _ensure_section(current_section, sections)
                    current_section.elements.append(DegradationBlock(
                        issue=DocumentIssue(
                            code=code,
                            message=f"Directive '{directive.name}' lexer issue: {code}.",
                            placement="block",
                        ),
                        fallback_text=directive.body,
                    ))
                if DIRECTIVE_SYNTAX_INVALID in directive.issues or DIRECTIVE_UNCLOSED in directive.issues:
                    body_lines = directive.body.split("\n")
                    sections, current_section = parse_block_lines(
                        body_lines, base_dir=base_dir, sections=sections, current_section=current_section
                    )
                continue

            handler = _DIRECTIVE_HANDLERS.get(directive.name)
            if handler is None:
                current_section = _ensure_section(current_section, sections)
                current_section.elements.append(DegradationBlock(
                    issue=DocumentIssue(
                        code=LONGFORM_DIRECTIVE_UNKNOWN,
                        message=f"Unknown directive '{directive.name}'.",
                        placement="block",
                    ),
                    fallback_text=directive.body,
                ))
                continue

            ctx = DirectiveContext(doc, seen_front_blocks, sections, current_section, base_dir)
            current_section = handler(directive, ctx, parse_block_lines, parse_inline)
            seen_front_blocks.update(ctx.seen_front_blocks)
            continue

    if not doc.title:
        all_lines: List[str] = []
        for s in sections:
            all_lines.append("#" * s.level + " " + s.heading if s.level > 0 else "")
            for elem in s.elements:
                if isinstance(elem, Paragraph):
                    all_lines.append(elem.plain_text)
        doc.title = detect_first_h1(all_lines)

    if not doc.title and sections and sections[0].has_heading:
        doc.title = sections[0].heading

    doc.sections = sections
    return doc


def _handle_abstract_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    if "abstract" in ctx.seen_front_blocks:
        ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
        ctx.current_section.elements.append(DegradationBlock(
            issue=DocumentIssue(
                code=LONGFORM_DUPLICATE_FRONT_BLOCK,
                message="Duplicate abstract block; only the first is used.",
                placement="block",
            ),
            fallback_text=directive.body,
        ))
        return ctx.current_section
    ctx.seen_front_blocks.add("abstract")

    paras: List[Paragraph] = []
    for elem in _collect_elements(parse_block_lines, directive.body.split("\n"), base_dir=ctx.base_dir):
        if isinstance(elem, Paragraph):
            paras.append(elem)
        elif isinstance(elem, ListBlock):
            for item in elem.items:
                text = "".join(s.text for s in item)
                if text:
                    paras.append(Paragraph(spans=parse_inline(text)))
    ctx.doc.abstract = AbstractBlock(paragraphs=paras)
    return ctx.current_section


def _handle_keywords_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    if "keywords" in ctx.seen_front_blocks:
        ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
        ctx.current_section.elements.append(DegradationBlock(
            issue=DocumentIssue(
                code=LONGFORM_DUPLICATE_FRONT_BLOCK,
                message="Duplicate keywords block; only the first is used.",
                placement="block",
            ),
            fallback_text=directive.body,
        ))
        return ctx.current_section
    ctx.seen_front_blocks.add("keywords")

    keywords: List[str] = []
    if "keywords" in directive.attributes:
        keywords = [
            k.strip()
            for k in directive.attributes["keywords"].replace(",", "\n").split("\n")
            if k.strip()
        ]
    else:
        for line in directive.body.splitlines():
            line = line.strip()
            if line:
                keywords.append(line)
    ctx.doc.keywords = KeywordsBlock(keywords=keywords)
    return ctx.current_section


def _handle_page_break_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
    ctx.current_section.elements.append(PageBreakBlock())
    return ctx.current_section


def _handle_table_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    table: Optional[TableBlock] = None
    for elem in _collect_elements(parse_block_lines, directive.body.split("\n"), base_dir=ctx.base_dir):
        if isinstance(elem, TableBlock):
            table = elem
            break
    if table is None:
        ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
        ctx.current_section.elements.append(DegradationBlock(
            issue=DocumentIssue(
                code="LONGFORM_TABLE_BODY_MISSING",
                message=":::table directive did not contain a Markdown table.",
                placement="block",
            ),
            fallback_text=directive.body,
        ))
        return ctx.current_section

    ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
    ctx.current_section.elements.append(SemanticTableBlock(
        identifier=directive.identifier,
        caption=directive.attributes.get("caption", ""),
        headers=table.headers,
        rows=table.rows,
        alignments=table.alignments,
    ))
    return ctx.current_section


def _handle_figure_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    images: List[ImageBlock] = []
    for elem in _collect_elements(parse_block_lines, directive.body.split("\n"), base_dir=ctx.base_dir):
        if isinstance(elem, ImageBlock):
            images.append(elem)

    ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
    ctx.current_section.elements.append(FigureBlock(
        identifier=directive.identifier,
        caption=directive.attributes.get("caption", ""),
        images=images,
        layout=directive.attributes.get("layout", "stack"),
    ))
    return ctx.current_section


def _handle_formula_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    source = directive.body.rstrip("\n")
    ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
    ctx.current_section.elements.append(FormulaBlock(
        identifier=directive.identifier,
        source=source,
    ))
    return ctx.current_section


def _handle_references_directive(
    directive: BlockDirective,
    ctx: DirectiveContext,
    parse_block_lines,
    parse_inline,
) -> Optional[Section]:
    entries: List[str] = []
    for line in directive.body.splitlines():
        line = line.rstrip()
        if line.strip():
            entries.append(line)
    ctx.current_section = _ensure_section(ctx.current_section, ctx.sections)
    ctx.current_section.elements.append(ReferenceListBlock(
        identifier=directive.identifier,
        entries=entries,
    ))
    return ctx.current_section


_DIRECTIVE_HANDLERS = {
    "abstract": _handle_abstract_directive,
    "keywords": _handle_keywords_directive,
    "page-break": _handle_page_break_directive,
    "table": _handle_table_directive,
    "figure": _handle_figure_directive,
    "formula": _handle_formula_directive,
    "references": _handle_references_directive,
}
