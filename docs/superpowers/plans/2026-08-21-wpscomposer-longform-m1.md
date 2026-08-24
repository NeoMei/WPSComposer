# WPSComposer Long-form M1 Semantic Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Build the 0.8.0 long-form semantic core from Markdown/frontmatter through a closed generation protocol v2 plan, with deterministic degradation and no live WPS startup.

**Architecture:** Add focused long-form modules for restricted YAML, block-directive lexing, semantic normalization, resources, IDs, formulas, layout policy, and plan construction. Extend the existing document model, parser, generation plan, and recording composer so legacy generation remains untouched while the new long-form path can be validated entirely offline.

**Tech Stack:** Python 3.9+, pytest, existing document model/parser/plan/recording modules, Pillow, SHA-256, NFC normalization, JSON-only plans.

## Global Constraints

- Implement only M1. Do not start Windows/macOS executors, real WPS, PDF quality gates, public generate routing changes, or M2-M5 behavior.
- Python API remains internal. Public authoring input is Markdown; the structured API is not a compatibility surface.
- No Python-object YAML, arbitrary YAML tags, anchor, alias, merge key, environment interpolation, template expansion, or network resource fetching.
- Visible text entering the semantic model is normalized to Unicode NFC. Fenced code is preserved and marked normalization none.
- Every local error becomes a deterministic inline/block/document degradation node. No unclosed or syntactically invalid directive may alter document structure.
- Generation plans must be JSON-compatible, deterministic for identical input, versioned, closed against unknown fields, resource-hash bound, and verifiable before WPS starts.
- Missing or invalid declared media becomes a degradation node. Staging filesystem failures remain fatal, but this milestone validates only the interface contract.
- uv.lock must remain untouched.
- Follow TDD for every behavior change: write failing tests, observe the expected failure, then implement.

## Task List

### Task 1: Restricted frontmatter parser

Files: create longform package, frontmatter_parser.py, and focused tests.

Interfaces: parse_frontmatter_document(text) returns values, issues, boundary, and body. Stable issues are FRONTMATTER_INVALID and FRONTMATTER_UNCLOSED.

- [ ] Write failing tests for a safe mapping, unclosed boundary, duplicate keys, tags/anchors/aliases/merge keys, object types, control characters, depth/key limits, and NFC output.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement bounded YAML 1.2 data-subset parsing through node composition only, never arbitrary object construction. Enforce 64 KiB boundary search, 256 total nodes, depth 8, duplicate keys, string-key mappings, scalar/list/mapping values, and fail-closed unknown tags.
- [ ] Run focused tests.
- [ ] Commit: Parse longform frontmatter safely.

### Task 2: Shared block-directive lexer

Files: create longform/directives.py and focused tests.

Interfaces: scan_block_directives(markdown) returns markdown regions and BlockDirective(name, identifier, attributes, body, start_line, issues). Stable issues are DIRECTIVE_SYNTAX_INVALID, DIRECTIVE_UNCLOSED, and NESTED_DIRECTIVE_UNSUPPORTED.

- [ ] Write failing tests for IDs, JSON-string attributes, unquoted tokens, duplicate keys, invalid escapes, nested directives, unclosed directives, code fences, indented code, block quotes, list contents, and NFC attributes.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement the shared lexer once, suppressing directive semantics outside top-level Markdown and preserving readable content on every syntax failure.
- [ ] Run focused tests.
- [ ] Commit: Lex longform block directives.

### Task 3: Unicode and bookmark identity helpers

Files: create unicode_text.py, bookmark_ids.py, and focused tests.

Interfaces: normalize_visible_text, display_units, shorten_display_units, contains_han, map_bookmarks, and BookmarkMapResult.

- [ ] Write failing tests for NFC, Han detection, display units, grapheme/ZWJ boundaries, deterministic bookmark ordering, ASCII bookmark format, and collision behavior.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement fixed Unicode helper tables needed by M1, without third-party data, and bookmark hashes using kind/NUL/id/NUL/attempt as specified.
- [ ] Run focused tests.
- [ ] Commit: Add longform Unicode and bookmark mapping.

### Task 4: Long-form semantic nodes and parser integration

Files: modify document_model.py and md_parser.py; create parser tests.

Interfaces: add DocumentIssue and long-form nodes AbstractBlock, KeywordsBlock, PageBreakBlock, SemanticTableBlock, FigureBlock, FormulaBlock, ReferenceListBlock, and DegradationBlock. parse_markdown(..., longform=True) produces them; longform=False preserves legacy behavior.

- [ ] Write failing tests for frontmatter metadata, abstract, keywords, explicit page breaks, invalid directives, duplicate front blocks, unknown directives, and legacy parser compatibility.
- [ ] Run focused tests and observe missing type/behavior failures.
- [ ] Extend the model and parser without changing legacy constructors or default behavior.
- [ ] Run new parser tests and existing parser tests.
- [ ] Commit: Model longform semantic blocks.

### Task 5: Semantic normalization, defaults, and reference resolution

Files: create longform/semantic.py and focused tests.

Interfaces: normalize_longform_document(doc) returns SemanticResult(document, config, references, bookmarks, issues) with deterministic to_json.

- [ ] Write failing tests for title consumption, explicit and automatic defaults, heading schemes, level gaps, header shortening, CONFIG_VALUE_INVALID, reference collection, unresolved references, bookmark mapping, and byte-stable snapshots.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement one-pass semantic normalization and validation.
- [ ] Run focused and all current longform tests.
- [ ] Commit: Normalize longform document semantics.

### Task 6: Resource preflight and formula subset contracts

Files: create longform/resources.py, longform/formula.py, and focused tests.

Interfaces: preflight_resources(nodes, base_dir) returns ResourcePreflight(resources, degradations, manifest); validate_formula_source(source) returns FormulaValidation.

- [ ] Write failing tests for missing resources, unsupported extensions, SVG manifest binding, valid raster/SVG hashes, redacted manifest output, formula subset acceptance, file/shell/macro rejection, and visible fallback text.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement path containment, media type checks, SHA-256 hashing, Pillow raster decode where available, SVG manifest compatibility, and conservative formula validation.
- [ ] Run focused tests.
- [ ] Commit: Preflight longform resources and formulas.

### Task 7: Long-form policy and generation protocol v2 plan

Files: create longform/policy.py and longform/plan.py; modify generation_plan.py and recording_composers.py; create plan tests.

Interfaces: build_longform_plan(semantic, preflight) returns a protocol v2 GenerationPlan. Add closed writer operations for begin/configure/title/front matter/numbering/TOC/index/caption/reference/formula/degradation/finalize.

- [ ] Write failing tests for deterministic identical plans, operation order, resource-hash binding, stable node IDs, no absolute paths, unknown-field rejection, missing-node rejection, and recording composer output.
- [ ] Run focused tests and observe missing modules/schema failures.
- [ ] Implement policy data, plan builder, protocol v2 schema, and recording methods while preserving protocol v1 behavior for existing tests.
- [ ] Run focused plan tests plus existing generation plan and recording composer tests.
- [ ] Commit: Build closed longform generation plans.

### Task 8: End-to-end M1 offline integration and deterministic snapshots

Files: create longform/pipeline.py and integration tests.

Interfaces: build_longform_generation(markdown, base_dir="") returns LongformBuild(document, semantic, preflight, plan, issues).

- [ ] Write failing tests for missing image degradation, stable semantic/plan snapshots, issue placement, and no WPS/process/artifact side effects.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement pipeline composition without importing platform executor modules or changing public generate.
- [ ] Run all longform M1 tests, then the full suite.
- [ ] Commit: Compose offline longform semantic pipeline.
