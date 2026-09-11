"""Read-only installed-font/PDF evidence; no Office calls or render-pass claim.

The negative gate is necessary, never sufficient for visual acceptance. It
rejects .notdef and broken Unicode mappings; manual layout/shaping review and
script coverage remain separate. No font substitutions or Unicode normalization.
Run with system python3 (local fontTools and PyMuPDF), output path must be new.
"""
from __future__ import annotations

import hashlib
import io
import json
from pathlib import Path


def missing_glyph_gate(text, pairs):
    if not isinstance(text, str) or not text or '\0' in text or '\ufffd' in text or not pairs:
        return False
    return all(
        isinstance(pair, (tuple, list)) and len(pair) == 2
        and type(pair[0]) is int and type(pair[1]) is int
        and 0 < pair[0] <= 0x10FFFF and pair[0] != 0xFFFD
        and not 0xD800 <= pair[0] <= 0xDFFF and pair[1] > 0
        for pair in pairs
    )


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def diagnose(pdf_path):
    import fitz
    from fontTools.ttLib import TTFont
    before = sha(pdf_path)
    report = {'scope': 'Missing-glyph negative evidence only; never full render acceptance',
              'pdf_sha256': before, 'installed_fonts': [], 'embedded_arial': []}
    for path in (Path('/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf'),
                 Path('/Applications/Microsoft Word.app/Contents/Resources/DFonts/arialbi.ttf')):
        with TTFont(path) as font:
            cmap = font.getBestCmap()
            report['installed_fonts'].append({'path': str(path), 'sha256': sha(path),
                'postscript_name': font['name'].getDebugName(6),
                'num_glyphs': font['maxp'].numGlyphs,
                'arabic_cmap': {char: cmap.get(ord(char)) for char in 'العربية'}})
    text = ''; pairs = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            text += page.get_text()
            for trace in page.get_texttrace():
                pairs.extend((char[0], char[1]) for char in trace['chars'])
        for entry in pdf.get_page_fonts(0, full=True):
            if 'Arial' not in entry[3]:
                continue
            name, ext, kind, data = pdf.extract_font(entry[0])
            with TTFont(io.BytesIO(data)) as font:
                report['embedded_arial'].append({'xref': entry[0], 'name': name,
                    'pdf_object': pdf.xref_object(entry[0]), 'glyph_order': font.getGlyphOrder(),
                    'cmaps': [{'platform': c.platformID, 'encoding': c.platEncID,
                               'format': c.format, 'mapping': c.cmap} for c in font['cmap'].tables]})
    report.update({'nul_count': text.count('\0'), 'replacement_count': text.count('\ufffd'),
                   'glyph_zero_count': sum(gid == 0 for _, gid in pairs),
                   'bad_trace_pairs': [p for p in pairs if p[0] in (0, 0xFFFD) or p[1] == 0],
                   'missing_glyph_gate': missing_glyph_gate(text, pairs),
                   'pdf_preserved': sha(pdf_path) == before})
    return report


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    report = diagnose(args.pdf)
    with args.output.open('x') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write('\n')
    print(json.dumps({key: report[key] for key in
        ('missing_glyph_gate', 'glyph_zero_count', 'nul_count', 'pdf_preserved')}))
