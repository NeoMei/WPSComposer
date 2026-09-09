"""Frozen v0.9.0 business inventory, not runtime capability discovery.

Implementation and evidence coverage are independent. Legacy unpinned direct
COM fallbacks do not establish support in the explicit Microsoft public API.
This module loads no hosts, inspects no installed applications and performs no I/O.
"""
from __future__ import annotations

BASELINE_COMMIT = '6dd3a00dff226096ad963cc68a65088865541c25'
RAW_OBJECT_MEMBERS = {
    'WriterComposer': ('doc', 'selection'),
    'SheetComposer': ('ws',),
    'SlideComposer': ('pres',),
    'BaseComposer': ('app', 'doc'),
}

_COMPOSERS = {'common': ('_base.py',
            'BaseComposer',
            ('app',
             'attach_active',
             'close',
             'doc',
             'export_pdf',
             'is_bound_to',
             'open_document',
             'save',
             'save_copy',
             'save_current',
             'supports_attached_save_copy')),
 'presentation': ('slide.py',
                  'SlideComposer',
                  ('add_blank_slide',
                   'add_bullets_slide',
                   'add_image',
                   'add_section_slide',
                   'add_shape',
                   'add_table',
                   'add_text_slide',
                   'add_textbox',
                   'add_title_slide',
                   'apply_design_preset',
                   'apply_format_patch',
                   'apply_layout_template',
                   'apply_structural_op',
                   'export_pdf',
                   'inspect_document',
                   'inspect_selection',
                   'pres',
                   'save_pptx',
                   'set_background_color',
                   'set_notes',
                   'set_slide_size',
                   'slide_count')),
 'spreadsheet': ('sheet.py',
                 'SheetComposer',
                 ('add_chart',
                  'add_sheet',
                  'add_title_row',
                  'apply_format_patch',
                  'apply_structural_op',
                  'autofit',
                  'conditional_format',
                  'export_pdf',
                  'freeze_panes',
                  'inspect_document',
                  'inspect_selection',
                  'merge_cells',
                  'rename_sheet',
                  'save_xlsx',
                  'select_sheet',
                  'set_borders',
                  'set_cell_style',
                  'set_column_width',
                  'set_formula',
                  'set_header_footer',
                  'set_range_style',
                  'set_row_height',
                  'write_cell',
                  'write_table',
                  'ws')),
 'writer': ('writer.py',
            'WriterComposer',
            ('add_bibliography_legacy',
             'add_bibliography_native',
             'add_bullet_list',
             'add_captioned_figure_fallback',
             'add_captioned_figure_native',
             'add_centered',
             'add_citation_paragraph',
             'add_code_lines',
             'add_cross_reference_fallback',
             'add_cross_reference_paragraph',
             'add_degradation_notice',
             'add_document_quality_notice',
             'add_equation_native',
             'add_equation_native_fallback',
             'add_equation_number_native',
             'add_floating_textbox',
             'add_heading',
             'add_heading2',
             'add_heading_level',
             'add_heading_level_native',
             'add_horizontal_line',
             'add_image',
             'add_image_block',
             'add_inline_degradation',
             'add_landscape_section_before_pending_heading',
             'add_merged_table',
             'add_numbered_list',
             'add_page_break',
             'add_paragraph',
             'add_paragraph_horizontal_line',
             'add_quality_notice_at_bookmark',
             'add_rich_paragraph',
             'add_section',
             'add_semantic_table_fallback',
             'add_semantic_table_native',
             'add_styled_paragraph',
             'add_table',
             'add_wordart',
             'apply_format_patch',
             'apply_heading_text_color',
             'apply_structural_op',
             'compact_terminal_paragraph',
             'configure_section',
             'degradation_checkpoint',
             'doc',
             'ensure_heading_styles',
             'ensure_styles',
             'export_pdf',
             'finalize_fields',
             'insert_caption_index_native',
             'insert_figure_index',
             'insert_table_index',
             'insert_toc',
             'insert_toc_with_styles',
             'inspect_document',
             'inspect_selection',
             'pagination_fragment_for_bookmark',
             'pagination_map_for_ranges',
             'refresh_bookmarks_and_references',
             'refresh_fields',
             'refresh_indexes',
             'repaginate_and_update_numbering',
             'repaginate_and_update_page_fields',
             'reserve_document_quality_anchor',
             'reset',
             'rollback_degradation_checkpoint',
             'save_docx',
             'selection',
             'set_columns',
             'set_document_metadata',
             'set_footer',
             'set_header',
             'set_header_footer',
             'set_margins',
             'set_orientation',
             'set_page_number_in_footer',
             'set_page_numbering',
             'set_page_role',
             'set_page_size',
             'snapshot_fields',
             'update_fields',
             'upsert_document_quality_notice'))}

_GENERATION = {'presentation': ('slide.add_blank',
                  'slide.add_bullets',
                  'slide.add_image',
                  'slide.add_section',
                  'slide.add_table',
                  'slide.add_title',
                  'slide.apply_preset',
                  'slide.reset',
                  'slide.set_size'),
 'spreadsheet': ('sheet.add',
                 'sheet.autofit',
                 'sheet.rename',
                 'sheet.reset',
                 'sheet.select',
                 'sheet.set_column_width',
                 'sheet.write_table'),
 'writer': ('writer.add_bibliography',
            'writer.add_captioned_figure',
            'writer.add_cross_reference',
            'writer.add_degradation_notice',
            'writer.add_document_quality_notice',
            'writer.add_equation',
            'writer.add_heading',
            'writer.add_horizontal_line',
            'writer.add_image',
            'writer.add_inline_degradation',
            'writer.add_list',
            'writer.add_page_break',
            'writer.add_paragraph',
            'writer.add_section',
            'writer.add_semantic_table',
            'writer.add_table',
            'writer.configure_front_matter',
            'writer.configure_page',
            'writer.configure_section',
            'writer.configure_toc_styles',
            'writer.ensure_styles',
            'writer.finalize_fields',
            'writer.insert_figure_index',
            'writer.insert_table_index',
            'writer.insert_toc',
            'writer.reserve_document_quality_anchor',
            'writer.reset',
            'writer.set_header_footer',
            'writer.set_page_number',
            'writer.set_page_numbering',
            'writer.set_page_role',
            'writer.update_fields')}

_PATCH_TARGETS = {'presentation': ('selection',
                  'presentation',
                  'slide:N',
                  'slide:N/shape:N',
                  'slide:N/shape:@id=N',
                  'slide:N/shape:@name=NAME',
                  'slide:N/shape:@id=N/paragraph:N',
                  'slide:N/shape:@id=N/paragraph:N/run:N',
                  'slide:N/shape:@id=N/table/cell:R,C',
                  'slide:N/shape:N/paragraph:N',
                  'slide:N/shape:N/paragraph:N/run:N',
                  'slide:N/shape:N/table/cell:R,C'),
 'spreadsheet': ('selection',
                 'sheet:N',
                 'sheet:N/cell:A1',
                 'sheet:N/range:A1:C20',
                 'sheet:N/shape:N',
                 'sheet:N/shape:@id=N',
                 'sheet:N/shape:@name=NAME',
                 'sheet:N/chart:N'),
 'writer': ('selection',
            'paragraph:N',
            'paragraph:@paraId=HEX',
            'range:S-E',
            'table:N/cell:R,C',
            'shape:N',
            'section:N')}

_INSERT_TYPES = {'presentation': ('slide', 'textbox', 'image'),
 'spreadsheet': ('row', 'column', 'sheet'),
 'writer': ('paragraph', 'heading', 'page_break', 'table', 'image', 'textbox')}

_MAC_WORD_OPERATIONS = ('add_bibliography',
 'add_captioned_figure',
 'add_cross_reference',
 'add_degradation_notice',
 'add_heading',
 'add_inline_degradation',
 'add_list',
 'add_page_break',
 'add_paragraph',
 'add_semantic_table',
 'configure_front_matter',
 'configure_page',
 'configure_section',
 'configure_toc_styles',
 'ensure_styles',
 'finalize_fields',
 'insert_figure_index',
 'insert_table_index',
 'insert_toc',
 'reserve_document_quality_anchor',
 'reset')



# Literal snapshots extracted from BASELINE_COMMIT, never from live backends.
_GENERATION_V1 = {'presentation': ('slide.add_blank',
                  'slide.add_bullets',
                  'slide.add_image',
                  'slide.add_section',
                  'slide.add_table',
                  'slide.add_title',
                  'slide.apply_preset',
                  'slide.reset',
                  'slide.set_size'),
 'spreadsheet': ('sheet.add',
                 'sheet.autofit',
                 'sheet.rename',
                 'sheet.reset',
                 'sheet.select',
                 'sheet.set_column_width',
                 'sheet.write_table'),
 'writer': ('writer.add_heading',
            'writer.add_horizontal_line',
            'writer.add_image',
            'writer.add_list',
            'writer.add_page_break',
            'writer.add_paragraph',
            'writer.add_section',
            'writer.add_table',
            'writer.configure_page',
            'writer.ensure_styles',
            'writer.insert_toc',
            'writer.reset',
            'writer.set_page_number',
            'writer.update_fields')}

_M5_WINDOWS_OPERATIONS = ('writer.add_bibliography',
 'writer.add_captioned_figure',
 'writer.add_cross_reference',
 'writer.add_degradation_notice',
 'writer.add_document_quality_notice',
 'writer.add_equation',
 'writer.add_heading',
 'writer.add_inline_degradation',
 'writer.add_list',
 'writer.add_page_break',
 'writer.add_paragraph',
 'writer.add_semantic_table',
 'writer.configure_front_matter',
 'writer.configure_page',
 'writer.configure_section',
 'writer.configure_toc_styles',
 'writer.ensure_styles',
 'writer.finalize_fields',
 'writer.insert_figure_index',
 'writer.insert_table_index',
 'writer.insert_toc',
 'writer.reserve_document_quality_anchor',
 'writer.reset',
 'writer.set_header_footer',
 'writer.set_page_numbering',
 'writer.set_page_role')

_INPUT_FORMATS = {'presentation': ('.dps',
                  '.dpt',
                  '.odp',
                  '.pot',
                  '.potm',
                  '.potx',
                  '.pps',
                  '.ppsm',
                  '.ppsx',
                  '.ppt',
                  '.pptm',
                  '.pptx'),
 'spreadsheet': ('.csv',
                 '.et',
                 '.ett',
                 '.htm',
                 '.html',
                 '.ods',
                 '.tsv',
                 '.xls',
                 '.xlsb',
                 '.xlsm',
                 '.xlsx',
                 '.xlt',
                 '.xltm',
                 '.xltx',
                 '.xml'),
 'writer': ('.doc',
            '.docm',
            '.docx',
            '.dot',
            '.dotm',
            '.dotx',
            '.htm',
            '.html',
            '.mht',
            '.mhtml',
            '.odt',
            '.rtf',
            '.txt',
            '.wps',
            '.wpt',
            '.xml')}

_CONVERSION_FORMATS = {'.doc': 'writer',
 '.docx': 'writer',
 '.ppt': 'presentation',
 '.pptx': 'presentation',
 '.xls': 'spreadsheet',
 '.xlsx': 'spreadsheet'}

_MAC_INSPECT_FORMATS = ('.doc',
 '.docm',
 '.docx',
 '.pps',
 '.ppsm',
 '.ppsx',
 '.ppt',
 '.pptm',
 '.pptx',
 '.xls',
 '.xlsm',
 '.xlsx')

_SAVE_FORMATS = {'presentation': ('.odp', '.potm', '.potx', '.ppsm', '.ppsx', '.ppt', '.pptm', '.pptx'),
 'spreadsheet': ('.csv', '.ods', '.tsv', '.xls', '.xlsb', '.xlsm', '.xlsx', '.xltm', '.xltx'),
 'writer': ('.doc',
            '.docm',
            '.docx',
            '.dotm',
            '.dotx',
            '.htm',
            '.html',
            '.mht',
            '.mhtml',
            '.odt',
            '.rtf',
            '.txt',
            '.xml',
            '.xps')}

_PATCH_DIMENSIONS = {'presentation': ('text',
                  'font',
                  'paragraph',
                  'geometry',
                  'fill',
                  'line',
                  'text_frame',
                  'name',
                  'shape_type',
                  'background',
                  'follow_master_background',
                  'page_setup',
                  'vertical_alignment'),
 'spreadsheet': ('value',
                 'formula',
                 'font',
                 'fill',
                 'line',
                 'geometry',
                 'number_format',
                 'horizontal_alignment',
                 'vertical_alignment',
                 'wrap_text',
                 'indent',
                 'row_height',
                 'column_width',
                 'borders',
                 'page_setup',
                 'name',
                 'chart_type',
                 'chart_title'),
 'writer': ('text',
            'font',
            'paragraph',
            'geometry',
            'fill',
            'line',
            'style',
            'wrap',
            'vertical_alignment',
            'page_setup',
            'columns')}

_PAGE_SETUP_PROPERTIES = {'presentation': ('slide_width', 'slide_height'),
 'spreadsheet': ('orientation',
                 'top_margin',
                 'bottom_margin',
                 'left_margin',
                 'right_margin',
                 'header_margin',
                 'footer_margin',
                 'paper_size',
                 'zoom',
                 'fit_to_pages_wide',
                 'fit_to_pages_tall',
                 'print_area',
                 'print_title_rows',
                 'print_title_columns'),
 'writer': ('orientation',
            'page_width',
            'page_height',
            'top_margin',
            'bottom_margin',
            'left_margin',
            'right_margin',
            'header_distance',
            'footer_distance',
            'gutter',
            'paper_size')}

_SHARED_PROPERTIES = {'fill': ('color', 'back_color', 'visible', 'transparency'),
 'font': ('name', 'size', 'bold', 'italic', 'underline', 'strikethrough', 'color'),
 'geometry': ('left', 'top', 'width', 'height', 'rotation'),
 'line': ('visible', 'weight', 'dash_style', 'transparency', 'color'),
 'paragraph': ('alignment',
               'left_indent',
               'right_indent',
               'first_line_indent',
               'space_before',
               'space_after',
               'line_spacing',
               'line_spacing_rule',
               'keep_together',
               'keep_with_next',
               'page_break_before',
               'widow_control')}

_PDF_METHODS = ('merge',
 'split',
 'extract_pages',
 'rotate',
 'extract_text',
 'page_count',
 'add_text_watermark')


def _state(implementation, note='', evidence=(), execution_routes=()):
    return {
        'implementation': implementation,
        'verification': 'representative_verified' if evidence else 'unverified',
        'evidence': list(evidence),
        'note': note,
        'execution_routes': list(execution_routes),
    }


def _row(identifier, component, operation, route, symbol):
    return {
        'id': identifier,
        'component': component,
        'operation': operation,
        'route': route,
        'source': symbol if symbol.startswith('macos/') else 'skills/WPSComposer/scripts/' + symbol,
        'baseline_commit': BASELINE_COMMIT,
        'required': True,
        'wps': {p: _state('unavailable') for p in ('win32', 'darwin')},
        'msoffice': {p: _state('unavailable') for p in ('win32', 'darwin')},
    }


def capability_records():
    """Return fresh JSON-safe records of baseline operations and their limits.

    A representative_verified row identifies bounded integration evidence,
    never exhaustive coverage of every argument combination. Detailed method
    and operation records remain unverified unless individually exercised.
    """
    rows = []
    evidence_root = 'docs/verification/msoffice-production/'
    representatives = {
        ('wps', 'darwin'): evidence_root + 'macos-wps/public-representative-all-schemes/report.json',
        ('wps', 'win32'): evidence_root + 'windows/round3-deb67ba/wps-representative-02/report.json',
        ('msoffice', 'darwin'): evidence_root + 'macos-word/public-representative-shared-outline/report.json',
        ('msoffice', 'win32'): evidence_root + 'windows/round4-shared-numbering/representative-04/report.json',
    }
    actions = ('generate', 'convert_to_pdf', 'inspect', 'edit', 'open_document', 'attach_active', 'apply_ops', 'apply_patches')
    for component in ('writer', 'spreadsheet', 'presentation'):
        for action in actions:
            filename = 'orchestrator.py' if action == 'generate' else 'conversion.py' if action == 'convert_to_pdf' else 'document_api.py'
            row = _row(f'public.{component}.{action}', component, action, 'public', filename + ':' + action)
            row['wps']['win32'] = _state('implemented', 'Existing platform contract; not all arguments natively verified.')
            if action in ('generate', 'convert_to_pdf', 'inspect'):
                row['wps']['darwin'] = _state('partial' if action == 'inspect' else 'implemented', 'Closed format/operation whitelist; Mac inspection covers only the separately enumerated source formats, with no active selection.')
            elif action == 'edit' and component == 'presentation':
                row['wps']['darwin'] = _state('partial', 'File-based PPTX set patches only; no structural operations or edit+PDF batch.')
            else:
                row['wps']['darwin'] = _state('unavailable', 'Public entry requires Windows COM in this baseline.')
            if component == 'writer' and action in ('generate', 'convert_to_pdf'):
                for engine in ('wps', 'msoffice'):
                    for platform in ('win32', 'darwin'):
                        partial = engine == 'msoffice' and action == 'generate' and platform == 'darwin'
                        row[engine][platform] = _state(
                            'partial' if partial else 'implemented',
                            'Microsoft conversion accepts DOC/DOCX only; Mac Word advanced-operation exclusions apply. Evidence is representative, not exhaustive.' if engine == 'msoffice' else 'Existing WPS public native representative; format/operation whitelist applies.',
                            (representatives[(engine, platform)],),
                        )
            rows.append(row)

    for component, (filename, cls, methods) in _COMPOSERS.items():
        for method in methods:
            if method in RAW_OBJECT_MEMBERS[cls]:
                continue
            row = _row(f'composer.{component}.{method}', component, method, 'direct_com', filename + ':' + cls + '.' + method)
            row['wps']['win32'] = _state('implemented', 'Direct COM interface; parameter-specific behavior is not individually certified.')
            row['msoffice']['win32'] = _state('legacy_unpinned', 'Legacy ProgID fallback exists but is not an explicit engine-selected public session or native parity proof.')
            row['wps']['darwin'] = _state('unavailable', 'Direct COM API is Windows-only; equivalent plan operations are listed separately.')
            rows.append(row)

    for component, operations in _GENERATION.items():
        for operation in operations:
            v1 = operation in _GENERATION_V1[component]
            m5 = operation in _M5_WINDOWS_OPERATIONS
            row = _row('plan.' + operation, component, operation, 'generation_plan', 'generation_plan.py:' + ('ALLOWED_OPERATIONS' if v1 else '_LONGFORM_WRITER_OPERATIONS'))
            row['protocol_versions'] = ([1, 2] if v1 else [2]) if component == 'writer' else [1]
            row['wps']['darwin'] = _state('implemented', 'Validation membership is broader than executor dispatch; these are only the listed execution routes.', execution_routes=(('legacy-v1',) if v1 else ()) + (('longform-v2-m5',) if m5 else ()))
            row['wps']['win32'] = _state('implemented' if m5 else 'partial', 'M5 dispatch exists only for the listed subset. V1 semantics are supplied by direct COM renderers, not a Windows V1 plan executor.', execution_routes=(('legacy-renderer-equivalent',) if v1 else ()) + (('longform-v2-m5',) if m5 else ()))
            if component == 'writer':
                if m5:
                    row['msoffice']['win32'] = _state('implemented', 'Native Word shared M5 executor; no legacy route. Argument/degradation policies apply and this row is not individually certified.', execution_routes=('longform-v2-m5',))
                else:
                    row['msoffice']['win32'] = _state('unavailable', 'V1 grammar concept is absent from Windows M5 dispatch; native Word explicitly rejects the public legacy layout route.')
                if operation.split('.', 1)[1] in _MAC_WORD_OPERATIONS:
                    row['msoffice']['darwin'] = _state('partial', 'Compiler accepts this M5 operation with argument-specific restrictions; no public legacy route; consult native-word guide.', execution_routes=('longform-v2-m5',))
            rows.append(row)

    for component, targets in _PATCH_TARGETS.items():
        for target in targets:
            row = _row(f'target.{component}.{target}', component, target, 'patch_target', 'document_api.py:PATCH_GRAMMAR')
            row['wps']['win32'] = _state('implemented', 'Supported property dimensions remain target-specific.')
            if component == 'presentation' and target not in ('selection', 'presentation'):
                row['wps']['darwin'] = _state('partial', 'Only the JSAPI PPTX formatting target subset is implemented; grammar membership alone does not prove support.')
            rows.append(row)
        for verb in ('set', 'insert', 'remove', 'move', 'clone'):
            row = _row(f'patch.{component}.{verb}', component, verb, 'patch_verb', 'document_api.py:ALL_OPS')
            row['wps']['win32'] = _state('implemented', 'Target-specific operations and attached-document atomic restrictions apply.')
            if verb == 'set' and component == 'presentation':
                row['wps']['darwin'] = _state('partial', 'PPTX file formatting subset only.')
            if verb == 'insert':
                row['insert_types'] = list(_INSERT_TYPES[component])
            rows.append(row)
        for element in _INSERT_TYPES[component]:
            row = _row(f'insert.{component}.{element}', component, element, 'insert_type', 'document_api.py:INSERT_TYPES')
            row['wps']['win32'] = _state('implemented', 'Native COM structural insertion; parent/position/props and attached atomic restrictions apply; row is not individually certified.')
            rows.append(row)

    _append_format_rows(rows)
    _append_property_rows(rows)
    for method in _PDF_METHODS:
        row = _row('composer.pdf.' + method, 'pdf', method, 'independent_pdf', 'pdf.py:PdfComposer.' + method)
        _shared_states(row, 'Independent shared PDF utility; Office engine is not involved. Optional pypdf/pdfplumber/reportlab dependencies apply; no native parity evidence implied.')
        rows.append(row)
    for filename, methods in (
        ('document_api.py', ('validate_op', 'validate_target', 'patch_grammar', 'snapshot_to_patches', 'snapshot_json', 'supported_formats', 'composer_for_path', 'composer_for_kind')),
        ('orchestrator.py', ('list_formats', 'list_available_presets')),
        ('md_parser.py', ('parse', 'parse_file')),
        ('plugins/__init__.py', ('list_plugins', 'register_plugin')),
    ):
        for method in methods:
            row = _row('utility.' + method, 'common', method, 'shared_utility', filename + ':' + method)
            _shared_states(row, 'Shared preparation/validation utility; does not establish a native application capability.')
            rows.append(row)
    return rows


def _shared_states(row, note):
    for engine in ('wps', 'msoffice'):
        for platform in ('win32', 'darwin'):
            row[engine][platform] = _state('implemented', note)


def _append_format_rows(rows):
    for component in ('writer', 'spreadsheet', 'presentation'):
        for source in ('markdown_file', 'markdown_text'):
            row = _row(f'input.{component}.generate.{source}', component, 'generate', 'generation_input', 'orchestrator.py:generate')
            row['format'] = source
            for platform in ('win32', 'darwin'):
                row['wps'][platform] = _state('implemented', 'UTF-8 Markdown file or source_is_text=True; source resources and generated output must satisfy the native route contract.')
                if component == 'writer':
                    row['msoffice'][platform] = _state('partial' if platform == 'darwin' else 'implemented', 'Public native M5 route only; Mac advanced-operation exclusions apply.')
            rows.append(row)
    for component, suffixes in _INPUT_FORMATS.items():
        for suffix in suffixes:
            for action in ('inspect', 'open_document', 'edit'):
                row = _row(f'format.{component}.{action}.{suffix}', component, action, 'source_format', 'document_api.py:supported_formats')
                row['format'] = suffix
                row['wps']['win32'] = _state('partial', 'Declared COM family routing; actual host format support is unverified. Shared XML/HTML suffixes default to Writer unless kind is explicit; edit output restrictions still apply.')
                if action == 'inspect' and suffix in _MAC_INSPECT_FORMATS:
                    row['wps']['darwin'] = _state('implemented', 'File inspection via the JSAPI INSPECTABLE whitelist; active selection is unavailable.')
                elif action == 'edit' and suffix == '.pptx':
                    row['wps']['darwin'] = _state('partial', 'PPTX set patches only; no structural edits, active binding, or edit+PDF batch.')
                rows.append(row)
        for suffix in _SAVE_FORMATS[component]:
            filename, cls, _ = _COMPOSERS[component]
            row = _row(f'format.{component}.save.{suffix}', component, 'save', 'save_format', filename + ':' + cls + '._formats_by_extension')
            row['format'] = suffix
            row['wps']['win32'] = _state('partial', 'Declared SaveAs constant; format-specific host execution remains unverified.')
            row['msoffice']['win32'] = _state('legacy_unpinned', 'Inherited direct COM fallback is not explicit native Microsoft session support.')
            rows.append(row)
    for suffix, component in _CONVERSION_FORMATS.items():
        row = _row(f'format.{component}.convert_to_pdf.{suffix}', component, 'convert_to_pdf', 'conversion_format', 'conversion.py:_COMPONENT_BY_SUFFIX')
        row['format'] = suffix
        for platform in ('win32', 'darwin'):
            row['wps'][platform] = _state('implemented', 'Public Office-to-PDF whitelist; this source format is not individually certified.')
            if component == 'writer':
                row['msoffice'][platform] = _state('implemented', 'Native Word DOC/DOCX conversion; source-specific native evidence is not inferred from generation.')
        rows.append(row)
    for component, formats in (('writer', ('.docx', '.pdf')), ('spreadsheet', ('.xlsx',)), ('presentation', ('.pptx',))):
        for fmt in formats:
            row = _row(f'format.{component}.generate.{fmt}', component, 'generate', 'generation_format', 'orchestrator.py:list_formats')
            row['format'] = fmt
            row['source_formats'] = ['Markdown file', 'Markdown text']
            for platform in ('win32', 'darwin'):
                row['wps'][platform] = _state('implemented', 'Public Markdown generation output; individual argument combinations remain unverified.')
                if component == 'writer':
                    row['msoffice'][platform] = _state('partial' if platform == 'darwin' else 'implemented', 'Native Word M5 route; legacy layout unavailable; Mac argument exclusions apply.')
            rows.append(row)


def _append_property_rows(rows):
    mac_properties = {
        'text', 'font', 'paragraph', 'geometry', 'fill', 'line', 'text_frame',
        'vertical_alignment', 'name', 'background', 'follow_master_background',
        *('font.' + x for x in _SHARED_PROPERTIES['font']),
        *('geometry.' + x for x in _SHARED_PROPERTIES['geometry']),
        *('paragraph.' + x for x in ('alignment', 'left_indent', 'first_line_indent', 'line_spacing', 'line_rule_within', 'space_before', 'space_after')),
        *('fill.' + x for x in ('color', 'visible', 'transparency')),
        *('background.' + x for x in ('color', 'visible', 'transparency')),
        *('line.' + x for x in ('color', 'visible', 'weight')),
    }
    text_frame = ('margin_left', 'margin_right', 'margin_top', 'margin_bottom', 'word_wrap', 'auto_size', 'vertical_anchor')
    mac_properties.update('text_frame.' + x for x in text_frame)
    for component, dimensions in _PATCH_DIMENSIONS.items():
        properties = set(dimensions)
        for dimension in dimensions:
            if dimension in _SHARED_PROPERTIES:
                properties.update(dimension + '.' + x for x in _SHARED_PROPERTIES[dimension])
        properties.update('page_setup.' + x for x in _PAGE_SETUP_PROPERTIES[component])
        if component == 'spreadsheet':
            properties.update('borders.*.' + x for x in ('style', 'weight', 'color'))
        if component == 'presentation':
            properties.add('paragraph.line_rule_within')
            properties.update('text_frame.' + x for x in text_frame)
            properties.update('background.' + x for x in _SHARED_PROPERTIES['fill'])
        filename, cls, _ = _COMPOSERS[component]
        for prop in sorted(properties):
            row = _row(f'property.{component}.{prop}', component, prop, 'patch_property', filename + ':' + cls + '.apply_format_patch')
            row['wps']['win32'] = _state('partial', 'Implemented for applicable targets only; COM property existence and accepted/rejected report must be checked. Not all target/property combinations are supported.')
            if prop == 'paragraph.line_rule_within':
                row['wps']['win32'] = _state('unavailable', 'Mac JSAPI-only paragraph property; absent from Windows apply_paragraph mapping.')
                row['source'] = 'macos/wps-jsapi-probe/addin/presentation.js:applyParagraphEdit'
            if component == 'presentation' and prop in mac_properties:
                row['wps']['darwin'] = _state('partial', 'JSAPI PPTX file formatting only; target-specific. Name is slide-only; unsupported keys can be omitted by the baseline JSAPI and are not certified.')
            rows.append(row)
