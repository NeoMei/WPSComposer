# Word story-range diagnostic matrix

This directory contains source-bound test fixtures and an opt-in Microsoft Word diagnostic. Nothing here is product code or native acceptance evidence.

## Source-only verification

From the `office-description` worktree root, use the clean development interpreter:

```bash
PYTHONDONTWRITEBYTECODE=1 /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-story-matrix/test_word_story_matrix.py
```

Regenerate the fixed assets only when intentionally changing their source contract:

```bash
PYTHONDONTWRITEBYTECODE=1 /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-story-matrix/generate_word_story_matrix.py \
  --out-dir .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-story-matrix
```

`preflight.json` is the authoritative source topology. `HASHES.json` binds the generator and generated assets.

## Parent-owned native commands

First independently confirm Word has no open documents. Run one fixture at a time, with a new output directory for every attempt:

```bash
PYTHONDONTWRITEBYTECODE=1 /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-story-matrix/word_story_matrix_probe.py \
  --execute-native \
  --source .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-story-matrix/absent.docx \
  --out build/word-story-matrix-20260910/absent-01
```

Repeat by replacing both fixture and output stem with `primary`, `inactive-first-even`, and `linked-multisection`. Do not reuse an output directory and do not continue after a retained/quarantined owned path.

The only successful terminal status is `STORY_MATRIX_DIAGNOSED`. It requires two byte-identical typed observation snapshots; unchanged saved/read-only/body End/body hash/list-template/paragraph/field/table/bookmark/section state around each of the six story-kind blocks; private/source byte equality; unchanged probe sources; acknowledged owned close; and final independent inventory `[]`.

Each story block reports one of these root outcomes: concrete `story-node` rows, `unresolved-property` with `native-missing` or `empty-class`, or `ordinary-property-error` with native number/text. A concrete node records class, story type, Start, End, exact content, section count, first section index, field count, and chain termination. Every concrete field records ordinal plus field type, code content/Start/End, result content/Start/End, and locked state. Missing and undefined native values remain explicit diagnostic outcomes; they are never accepted as empty content or zero fields.

The diagnostic deliberately avoids `text object of header` and `text object of footer`. It does not prove a production replacement, full quality snapshot, save/export, UI, WPS, Windows, or cross-platform parity.
