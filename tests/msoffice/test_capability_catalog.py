from __future__ import annotations

import ast
import json
from pathlib import Path
import subprocess
import sys

from skills.WPSComposer.scripts.capability_catalog import (
    BASELINE_COMMIT, RAW_OBJECT_MEMBERS, capability_records,
)

ROOT = Path(__file__).resolve().parents[2]


def frozen_tree(filename):
    source = subprocess.run(
        ['git', 'show', BASELINE_COMMIT + ':skills/WPSComposer/scripts/' + filename],
        cwd=ROOT, check=True, capture_output=True, text=True,
    ).stdout
    return ast.parse(source)


def assignment(tree, name):
    return next(node.value for node in tree.body if isinstance(node, ast.Assign)
                and any(isinstance(t, ast.Name) and t.id == name for t in node.targets))


def index():
    return {row['id']: row for row in capability_records()}


def test_records_are_unique_frozen_and_source_bound():
    rows = capability_records()
    assert rows and len(rows) == len(index())
    assert BASELINE_COMMIT == '6dd3a00dff226096ad963cc68a65088865541c25'
    for row in rows:
        assert row['baseline_commit'] == BASELINE_COMMIT
        assert (ROOT / row['source'].split(':')[0]).is_file()
        for engine in ('wps', 'msoffice'):
            assert set(row[engine]) == {'win32', 'darwin'}
            for status in row[engine].values():
                assert status['implementation'] in {'implemented', 'partial', 'unavailable', 'legacy_unpinned'}
                assert status['verification'] in {'unverified', 'representative_verified'}
                if status['verification'] == 'representative_verified':
                    assert status['evidence']
                    assert all((ROOT / p).is_file() for p in status['evidence'])


def test_baseline_does_not_claim_microsoft_excel_or_powerpoint_routes():
    rows = index()
    for component in ('spreadsheet', 'presentation'):
        for action in ('generate', 'convert_to_pdf', 'inspect', 'edit', 'attach_active'):
            for platform in ('win32', 'darwin'):
                status = rows[f'public.{component}.{action}']['msoffice'][platform]
                assert status['implementation'] == 'unavailable'
                assert status['verification'] == 'unverified'
    assert rows['public.writer.generate']['msoffice']['darwin']['implementation'] == 'partial'
    assert rows['public.writer.generate']['msoffice']['darwin']['verification'] == 'representative_verified'


def test_macos_wps_editing_limits_are_not_overstated():
    rows = index()
    assert rows['public.presentation.edit']['wps']['darwin']['implementation'] == 'partial'
    for component in ('writer', 'spreadsheet'):
        assert rows[f'public.{component}.edit']['wps']['darwin']['implementation'] == 'unavailable'
    for component in ('writer', 'spreadsheet', 'presentation'):
        assert rows[f'public.{component}.attach_active']['wps']['darwin']['implementation'] == 'unavailable'
        assert rows[f'public.{component}.inspect']['wps']['darwin']['implementation'] == 'partial'
        assert rows[f'patch.{component}.insert']['wps']['darwin']['implementation'] == 'unavailable'


def test_catalog_covers_every_declared_semantic_composer_member():
    rows = index()
    for filename, cls, component in (
        ('writer.py', 'WriterComposer', 'writer'),
        ('sheet.py', 'SheetComposer', 'spreadsheet'),
        ('slide.py', 'SlideComposer', 'presentation'),
        ('_base.py', 'BaseComposer', 'common'),
    ):
        tree = frozen_tree(filename)
        klass = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == cls)
        names = {node.name for node in klass.body if isinstance(node, ast.FunctionDef) and not node.name.startswith('_')}
        for name in names - set(RAW_OBJECT_MEMBERS[cls]):
            assert f'composer.{component}.{name}' in rows, (cls, name)
        for name in RAW_OBJECT_MEMBERS[cls]:
            assert f'composer.{component}.{name}' not in rows


def test_records_are_defensive_copies_and_json_serializable():
    first = capability_records()
    expected = json.dumps(first, sort_keys=True)
    first[0]['wps']['darwin']['implementation'] = 'fake'
    assert json.dumps(capability_records(), sort_keys=True) == expected


def test_catalog_import_does_not_load_native_backends():
    code = "import sys; from skills.WPSComposer.scripts.capability_catalog import capability_records; capability_records(); assert not any(x.startswith(('win32com', 'skills.WPSComposer.scripts.msoffice.', 'skills.WPSComposer.scripts.macos_probe.')) for x in sys.modules)"
    result = subprocess.run([sys.executable, '-c', code], cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


def test_all_pdf_business_methods_have_independent_four_platform_rows():
    rows = index()
    klass = next(n for n in frozen_tree('pdf.py').body if isinstance(n, ast.ClassDef) and n.name == 'PdfComposer')
    methods = {n.name for n in klass.body if isinstance(n, ast.FunctionDef) and not n.name.startswith('_')}
    assert methods == {'merge', 'split', 'extract_pages', 'rotate', 'extract_text', 'page_count', 'add_text_watermark'}
    for method in methods:
        row = rows['composer.pdf.' + method]
        assert row['route'] == 'independent_pdf'
        for engine in ('wps', 'msoffice'):
            for status in row[engine].values():
                assert status['implementation'] == 'implemented'
                assert status['verification'] == 'unverified'


def test_format_rows_cover_declared_input_conversion_and_save_lists():
    rows = index()
    api = frozen_tree('document_api.py')
    for component, symbol in [('writer', 'WRITER_EXTENSIONS'), ('spreadsheet', 'SHEET_EXTENSIONS'), ('presentation', 'SLIDE_EXTENSIONS')]:
        for suffix in ast.literal_eval(assignment(api, symbol)):
            assert f'format.{component}.inspect.{suffix}' in rows
            assert f'format.{component}.open_document.{suffix}' in rows
            assert f'format.{component}.edit.{suffix}' in rows
    for suffix, component in ast.literal_eval(assignment(frozen_tree('conversion.py'), '_COMPONENT_BY_SUFFIX')).items():
        assert f'format.{component}.convert_to_pdf.{suffix}' in rows
    for component, filename in [('writer', 'writer.py'), ('spreadsheet', 'sheet.py'), ('presentation', 'slide.py')]:
        cls = next(n for n in frozen_tree(filename).body if isinstance(n, ast.ClassDef) and n.name == {'writer': 'WriterComposer', 'spreadsheet': 'SheetComposer', 'presentation': 'SlideComposer'}[component])
        formats = assignment(cls, '_formats_by_extension')
        for key in formats.keys:
            assert f'format.{component}.save.{ast.literal_eval(key)}' in rows
    assert rows['format.writer.inspect..odt']['wps']['darwin']['implementation'] == 'unavailable'
    assert rows['format.presentation.edit..pptm']['wps']['darwin']['implementation'] == 'unavailable'
    assert rows['format.writer.convert_to_pdf..doc']['msoffice']['darwin']['implementation'] == 'implemented'


def test_v1_only_operations_do_not_claim_m5_execution():
    rows = index()
    for operation in ('add_table', 'add_image', 'add_section', 'add_horizontal_line', 'set_page_number', 'update_fields'):
        row = rows['plan.writer.' + operation]
        assert row['protocol_versions'] == [1, 2]  # validation union is broader than dispatch
        assert row['msoffice']['win32']['implementation'] == 'unavailable'
        assert row['msoffice']['darwin']['implementation'] == 'unavailable'
        assert row['wps']['darwin']['execution_routes'] == ['legacy-v1']
    assert rows['plan.writer.configure_section']['protocol_versions'] == [2]
    assert rows['plan.writer.configure_section']['msoffice']['win32']['execution_routes'] == ['longform-v2-m5']


def test_plan_union_and_targets_match_frozen_contract():
    rows = index()
    plan = frozen_tree('generation_plan.py')
    union = ast.literal_eval(assignment(plan, 'ALLOWED_OPERATIONS'))
    union['writer'] |= ast.literal_eval(assignment(plan, '_LONGFORM_WRITER_OPERATIONS').args[0])
    assert {r['operation'] for r in rows.values() if r['route'] == 'generation_plan'} == set().union(*union.values())
    grammar = ast.literal_eval(assignment(frozen_tree('document_api.py'), 'PATCH_GRAMMAR'))
    for kind, component in [('writer', 'writer'), ('sheet', 'spreadsheet'), ('slide', 'presentation')]:
        assert {r['operation'] for r in rows.values() if r['route'] == 'patch_target' and r['component'] == component} == {entry[0] for entry in grammar[kind]}
    assert rows['target.presentation.presentation']['wps']['darwin']['implementation'] == 'unavailable'
    assert rows['target.presentation.slide:N/shape:@id=N/paragraph:N/run:N']['wps']['darwin']['implementation'] == 'partial'


def test_all_patch_dimensions_and_nested_properties_are_enumerated():
    rows = index()
    for component, filename in [('writer', 'writer.py'), ('spreadsheet', 'sheet.py'), ('presentation', 'slide.py')]:
        cls = next(n for n in frozen_tree(filename).body if isinstance(n, ast.ClassDef) and n.name == {'writer': 'WriterComposer', 'spreadsheet': 'SheetComposer', 'presentation': 'SlideComposer'}[component])
        patch = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == 'apply_format_patch')
        for arg in patch.args.kwonlyargs:
            assert f'property.{component}.{arg.arg}' in rows
    for component in ('writer', 'spreadsheet', 'presentation'):
        for key in ('font.name', 'font.color', 'geometry.rotation', 'line.dash_style', 'fill.back_color'):
            assert f'property.{component}.{key}' in rows
    for key in ('page_setup.gutter', 'paragraph.widow_control'):
        assert f'property.writer.{key}' in rows
    for key in ('borders.*.color', 'page_setup.print_title_columns'):
        assert f'property.spreadsheet.{key}' in rows
    assert rows['property.presentation.paragraph.line_rule_within']['wps']['win32']['implementation'] == 'unavailable'
    assert rows['property.presentation.paragraph.line_rule_within']['wps']['darwin']['implementation'] == 'partial'
    for key in ('line.dash_style', 'fill.back_color', 'page_setup', 'shape_type', 'paragraph.right_indent'):
        assert rows['property.presentation.' + key]['wps']['darwin']['implementation'] == 'unavailable'


def test_checked_in_baseline_exactly_matches_data_and_retains_gaps():
    baseline = json.loads((ROOT / 'docs/verification/microsoft-parity/baseline.json').read_text())
    assert baseline['baseline_version'] == 'v0.9.0'
    assert baseline['baseline_commit'] == BASELINE_COMMIT
    assert baseline['records'] == capability_records()
    assert baseline['parity_complete'] is False


def test_insert_types_and_public_input_formats_are_individual_capabilities():
    rows = index()
    types = ast.literal_eval(assignment(frozen_tree('document_api.py'), 'INSERT_TYPES'))
    for kind, component in [('writer', 'writer'), ('sheet', 'spreadsheet'), ('slide', 'presentation')]:
        for element in types[kind]:
            row = rows[f'insert.{component}.{element}']
            assert row['wps']['win32']['implementation'] == 'implemented'
            assert row['wps']['darwin']['implementation'] == 'unavailable'
            assert all(s['implementation'] == 'unavailable' for s in row['msoffice'].values())
        for source in ('markdown_file', 'markdown_text'):
            assert f'input.{component}.generate.{source}' in rows


def test_nested_formatting_mapping_coverage_comes_from_frozen_implementation():
    rows = index()
    tree = frozen_tree('formatting.py')
    for function, dimension in [('apply_font', 'font'), ('apply_geometry', 'geometry'), ('apply_fill', 'fill'), ('apply_line', 'line'), ('apply_paragraph', 'paragraph')]:
        method = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == function)
        keys = set()
        for node in ast.walk(method):
            if isinstance(node, ast.Assign) and any(isinstance(n, ast.Name) and n.id == 'mapping' for n in node.targets):
                keys.update(ast.literal_eval(node.value))
            if isinstance(node, ast.Compare) and isinstance(node.left, ast.Name) and node.left.id == 'key':
                keys.update(n.value for n in node.comparators if isinstance(n, ast.Constant) and isinstance(n.value, str))
        for component in ('writer', 'spreadsheet', 'presentation'):
            if dimension == 'paragraph' and component == 'spreadsheet':
                continue
            assert keys <= {r['operation'].split('.', 1)[1] for r in rows.values()
                            if r['component'] == component and r['route'] == 'patch_property'
                            and r['operation'].startswith(dimension + '.')}
