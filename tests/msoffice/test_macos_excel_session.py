from __future__ import annotations

import importlib
from pathlib import Path
import subprocess
import time

import pytest


def module():
    return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_excel_session')


def session(tmp_path, read_only=False):
    mod = module()
    obj = mod.MacExcelSession.__new__(mod.MacExcelSession)
    obj._source = tmp_path / 'source.xlsx'
    obj._native = tmp_path / 'owned.xlsx'
    obj._job = tmp_path
    obj._deadline = time.monotonic() + 600
    obj._read_only = read_only
    obj._closed = False
    obj._failed = False
    obj._attached = False
    obj._counter = 0
    obj._lock = None
    obj._execute = subprocess.run
    return obj


def test_target_parser_validates_excel_bounds_and_stable_shape_names():
    mod = module()
    assert mod.parse_target('sheet:2/range:$A$1:$C$8') == (2, 'range', 'A1:C8')
    assert mod.parse_target('sheet:1/shape:@name=Revenue "chart"') == (1, 'shape_name', 'Revenue "chart"')
    for target in ('sheet:0', 'sheet:1/cell:XFE1', 'sheet:1/range:C3:A1', 'sheet:1/cell:A1048577', 'sheet:1/cell:A1"\nquit'):
        with pytest.raises(ValueError):
            mod.parse_target(target)


def test_patch_compile_rejects_unknown_before_any_native_execution(tmp_path):
    obj = session(tmp_path)
    obj._run = lambda *a, **k: pytest.fail('no native execution')
    with pytest.raises(ValueError):
        obj.apply_format_patch('sheet:1/cell:A1', arbitrary='bad')
    with pytest.raises(ValueError):
        obj.apply_format_patch('sheet:1/cell:A1', font={'size': -1})


def test_read_only_patch_rejected_before_execution(tmp_path):
    obj = session(tmp_path, read_only=True)
    obj._run = lambda *a, **k: pytest.fail('no native execution')
    with pytest.raises(PermissionError):
        obj.apply_format_patch('sheet:1/cell:A1', value=4)


def test_patch_acceptance_is_per_property_and_quoted(tmp_path):
    obj = session(tmp_path)
    scripts = []
    obj._run = lambda source, **k: scripts.append(source) or {'accepted': ['value', 'font.bold'], 'rejected': []}
    result = obj.apply_format_patch('sheet:1/cell:A1', value='"line\nnext', font={'bold': True})
    assert result['accepted'] == ['value', 'font.bold']
    assert '\\"line" & linefeed & "next' in scripts[0]
    assert 'set bold of font object' in scripts[0]


def test_atomic_unknown_format_keys_cannot_silently_pass(tmp_path):
    obj = session(tmp_path)
    obj._run = lambda *a, **k: pytest.fail('must validate entire patch first')
    with pytest.raises(ValueError):
        obj.apply_format_patch('sheet:1/cell:A1', value=4, fill={'unexpected': 1})


def test_structural_copy_cut_only_for_explicit_ops_and_discloses_clipboard(tmp_path):
    obj=session(tmp_path);scripts=[]
    obj._run=lambda source,**k:scripts.append(source) or {'clipboard_changed':True}
    for verb in ('move','clone'):
        result=obj.apply_structural_op({'op':verb,'target':'sheet:1/cell:A2','to':{'index':4}})
        assert result['clipboard_changed'] is True
    assert 'cut range' in scripts[0] and 'copy range' in scripts[1]
    assert 'insert into range' in scripts[0]


def test_sheet_move_clone_rebinds_after_structural_change(tmp_path):
    obj=session(tmp_path);scripts=[]
    obj._run=lambda source,**k:scripts.append(source) or {'path':'sheet:2'}
    obj.apply_structural_op({'op':'move','target':'sheet:1','to':{'after':2}})
    obj.apply_structural_op({'op':'clone','target':'sheet:1','to':{'before':2}})
    assert 'move obj to after worksheet 2 of ownedBook' in scripts[0]
    assert 'copy worksheet obj before worksheet 2 of ownedBook' in scripts[1]
    assert 'beforeSheetCount' in scripts[1]


def test_inspection_limit_is_enforced_without_native_calls(tmp_path):
    obj = session(tmp_path)
    obj._run = lambda *a, **k: pytest.fail('must validate options')
    for limit in (-1, True, 100001):
        with pytest.raises(ValueError):
            obj.inspect_document(max_cells=limit)


def test_attached_session_never_closes_or_rebinds_document(tmp_path):
    obj=session(tmp_path);obj._attached=True
    obj._run=lambda *a,**k:pytest.fail('attached close must not send close event')
    with pytest.raises(NotImplementedError):obj.save(tmp_path/'copy.xlsx')
    with pytest.raises(NotImplementedError):obj.export_pdf(tmp_path/'copy.pdf')
    obj.close()
    assert obj._closed


def test_binding_and_context_close(tmp_path):
    obj = session(tmp_path)
    assert obj.is_bound_to(obj._native)
    assert not obj.is_bound_to(obj._source)
    calls=[]
    obj.close=lambda save_changes=False: calls.append(save_changes)
    with obj as entered:
        assert entered is obj
    assert calls==[False]
    assert obj.supports_attached_save_copy() is False


def test_timeout_quarantines_and_preserves_raw_step(tmp_path):
    obj=session(tmp_path)
    quarantines=[]
    class Lock:
        def quarantine(self,detail): quarantines.append(detail)
    obj._lock=Lock()
    def execute(*args,**kwargs):
        raise subprocess.TimeoutExpired(args[0],kwargs['timeout'],output=b'partial',stderr=b'late AppleEvent')
    obj._execute=execute
    with pytest.raises(RuntimeError):
        obj._run('return "ok"')
    assert obj._failed is True and quarantines
    assert list(tmp_path.glob('step-*.applescript'))
    assert any(p.read_text()=='late AppleEvent' for p in tmp_path.glob('*.stderr.log'))


def test_close_failed_session_does_not_send_another_event(tmp_path):
    obj=session(tmp_path)
    obj._failed=True
    obj._run=lambda *a,**k: pytest.fail('late events may still run')
    obj.close()
    assert obj._closed is True


def test_save_never_overwrites_source_or_existing_destination(tmp_path):
    obj=session(tmp_path)
    obj._source.write_bytes(b'keep source')
    obj._run=lambda *a,**k: pytest.fail('must reject before native call')
    with pytest.raises(ValueError): obj.save(obj._source)
    destination=tmp_path/'existing.xlsx';destination.write_bytes(b'keep destination')
    with pytest.raises(FileExistsError): obj.save(destination)


def test_inspection_cell_references_keep_worksheet_parent(tmp_path):
    obj=session(tmp_path)
    scripts=[]
    obj._run=lambda body,**kwargs:scripts.append(body) or {}
    obj.inspect_document(max_cells=10)
    assert 'cell (firstRow + r - 1) of column (firstCol + c - 1) of ws' in scripts[0]
    assert 'cell r of column c of usedRng' not in scripts[0]


def test_expired_session_quarantines_owned_document_before_releasing_lock(tmp_path):
    obj=session(tmp_path);obj._deadline=time.monotonic()-1;records=[]
    class Lock:
        def quarantine(self,detail):records.append(detail)
        def close(self):pass
    obj._lock=Lock()
    obj._execute=lambda *a,**k:pytest.fail('expired job must not launch another event')
    with pytest.raises(RuntimeError):obj.close()
    assert obj._closed and obj._failed and records


def test_shape_fill_and_line_compilation_has_native_object_properties(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {'accepted':[],'rejected':[]}
    obj.apply_format_patch('sheet:1/shape:@name=Revenue',fill={'color':'#F0F0F0','transparency':0.25},line={'color':'#123456','weight':2,'visible':True})
    assert 'fore color of fill format of obj' in scripts[0]
    assert 'weight of line format of obj' in scripts[0]


def test_stable_shape_name_requires_unique_owned_match(tmp_path):
    obj=session(tmp_path)
    source,kind=obj._target_script('sheet:1/shape:@name=Stable')
    assert kind=='shape_name'
    assert 'shapeMatches' in source and 'is not 1' in source


def test_color_snapshot_normalizes_back_color():
    assert module()._normalize({'color':[1,2,3],'back_color':[4,5,6]})=={'color':'#010203','back_color':'#040506'}


def test_attached_preflight_rejects_save_copy_and_unsaved_path_before_mutation(tmp_path):
    obj=session(tmp_path);obj._attached=True
    obj._run=lambda *a,**k:pytest.fail('filesystem preflight must reject unsaved path')
    with pytest.raises(NotImplementedError):obj.preflight_save(tmp_path/'copy.xlsx')
    with pytest.raises(ValueError):obj.preflight_save()


def test_attached_inspection_restores_owned_sheet_and_selection(tmp_path):
    obj=session(tmp_path);obj._attached=True;scripts=[]
    obj._run=lambda body,**k:scripts.append(body) or {}
    obj.inspect_document(max_cells=2)
    assert 'originalSheetName' in scripts[0]
    assert 'select range originalSelectionAddress of worksheet originalSheetName of ownedBook' in scripts[0]


@pytest.mark.parametrize('patch', [{'formula':'=RTD("x",,"y")'}, {'value':'=WEBSERVICE("https://example.com")'}, {'value':[['=RTD("x",,"y")']]}])
def test_unsafe_formula_patch_rejected_before_native_event(tmp_path,patch):
    obj=session(tmp_path)
    obj._run=lambda *a,**k:pytest.fail('unsafe formula must be rejected before AppleEvents')
    with pytest.raises(ValueError):obj.apply_format_patch('sheet:1/cell:A1',**patch)


def test_unsafe_insert_values_rejected_before_native_event(tmp_path):
    obj=session(tmp_path)
    obj._run=lambda *a,**k:pytest.fail('unsafe formula must be rejected before AppleEvents')
    with pytest.raises(ValueError):obj.apply_structural_op({'op':'insert','parent':'sheet:1','type':'row','position':{'index':2},'props':{'values':['=RTD("x",,"y")']}})


def test_direct_business_methods_use_selected_sheet_and_validate_before_mutation(tmp_path):
    obj=session(tmp_path);scripts=[]
    obj._run=lambda body,**k:scripts.append(body) or {'path':'sheet:2'}
    obj.select_sheet(2)
    obj.write_cell(3,2,10)
    obj.set_formula(4,2,'=B3*2')
    obj.merge_cells('A6:C6')
    obj.set_header_footer(left='Native',center='&P',right='&D')
    assert all('worksheet 2 of ownedBook' in s for s in scripts[1:])
    assert any('merge obj' in s for s in scripts)
    assert any('left header' in s for s in scripts)
    before=len(scripts)
    with pytest.raises(ValueError):obj.set_formula(1,1,'=RTD("x",,"y")')
    with pytest.raises(ValueError):obj.write_cell(0,1,4)
    assert len(scripts)==before


def test_direct_chart_freeze_and_conditional_format_compile_owned_native_objects(tmp_path):
    obj=session(tmp_path);scripts=[]
    obj._run=lambda body,**k:scripts.append(body) or {}
    obj.add_chart(source_range='A1:B4',title='Native revenue')
    obj.freeze_panes('B2')
    obj.conditional_format('B2:B4',formula='15')
    assert 'make new chart object' in scripts[0] and 'set source data' in scripts[0]
    assert 'freeze panes of ownedWindow' in scripts[1]
    assert 'format condition' in scripts[2]


def test_direct_table_prevalidates_all_values_before_first_write(tmp_path):
    obj=session(tmp_path)
    obj._run=lambda *a,**k:pytest.fail('entire table must be validated')
    with pytest.raises(ValueError):obj.write_table(1,1,[['safe'],['=RTD("x",,"y")']])


@pytest.mark.parametrize('method,suffix',[('save','.xlsx'),('export_pdf','.pdf')])
def test_direct_publication_uses_atomic_transport_with_session_deadline(tmp_path,monkeypatch,method,suffix):
    mod=module();obj=session(tmp_path);calls=[]
    obj._native.write_bytes(b'native bytes');obj._run=lambda *a,**k:{}
    monkeypatch.setattr(mod,'validate_before_deadline',lambda *a,**k:None)
    def publish(staged,destination,**options):
        calls.append((staged,destination,options));raise TimeoutError('injected copy deadline')
    monkeypatch.setattr(mod,'publish_artifact',publish,raising=False)
    destination=tmp_path/('public'+suffix)
    with pytest.raises(TimeoutError):getattr(obj,method)(destination)
    assert not destination.exists()
    assert calls and calls[0][2]['deadline']==obj._deadline and calls[0][2]['overwrite'] is False


def test_attached_inspection_restores_selection_in_native_error_handler(tmp_path):
    obj=session(tmp_path);obj._attached=True;scripts=[]
    obj._run=lambda body,**k:scripts.append(body) or {}
    obj.inspect_document(max_cells=2)
    assert 'on error inspectionError number inspectionNumber' in scripts[0]
    assert scripts[0].count('select range originalSelectionAddress')==2


def test_existing_worksheet_delete_rejects_before_any_native_mutation(tmp_path):
    obj=session(tmp_path);obj._run=lambda *a,**k:pytest.fail('existing sheet may prompt')
    with pytest.raises(NotImplementedError):obj.apply_structural_op({'op':'remove','target':'sheet:2'})


def test_direct_business_colors_accept_bgr_integer_contract():
    assert module()._color(0x563412)=='{18, 52, 86}'
    with pytest.raises(ValueError):module()._color(True)


def test_chart_and_condition_creation_use_verified_native_references(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.add_chart(source_range='A1:B4')
    obj.conditional_format('B2:B4')
    assert 'set nativeChart to chart of co' in scripts[0]
    assert 'set source data nativeChart source range "A1:B4" of ws' in scripts[0]
    assert 'make new format condition at obj with properties' in scripts[1]


def test_merge_passes_bound_range_object_directly(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.merge_cells('A1:C1');obj.add_title_row('A3:C3','Title')
    assert all('\nmerge obj\n' in body for body in scripts)


@pytest.mark.parametrize('verb',['move','clone','remove'])
def test_structural_multi_cell_range_uses_only_first_row_or_column(tmp_path,verb):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.apply_structural_op({'op':verb,'target':'sheet:1/range:B2:D5','axis':'row','to':{'index':8}})
    assert 'set obj to range "B2" of ws' in scripts[0]


def test_conditional_format_deletes_existing_rules_individually(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.conditional_format('B2:B3')
    assert 'repeat (count of format conditions of obj) times' in scripts[0]
    assert 'delete format condition 1 of obj' in scripts[0]


def test_inspection_reads_native_freeze_state_and_headers(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.inspect_document(max_cells=1)
    assert 'freeze panes of window 1 of ownedBook' in scripts[0]
    assert 'left header of pp' in scripts[0]
    assert 'freezeJSON &' in scripts[0]


def test_selected_sheet_tracks_own_sheet_move_and_clone(tmp_path):
    obj=session(tmp_path);obj._sheet_index=2
    obj._run=lambda *a,**k:{'path':'sheet:1'}
    obj.apply_structural_op({'op':'move','target':'sheet:2','to':{'before':1}})
    assert obj._sheet_index==1
    obj._run=lambda *a,**k:{'path':'sheet:1'}
    obj.apply_structural_op({'op':'clone','target':'sheet:2','to':{'before':1}})
    assert obj._sheet_index==2


def test_explicit_row_clone_finishes_its_native_cut_copy_mode(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.apply_structural_op({'op':'clone','target':'sheet:1/cell:A2','to':{'index':4}})
    assert 'set cut copy mode to false' in scripts[0]


def test_new_document_creates_native_workbook_in_private_locked_job(tmp_path,monkeypatch):
    mod=module();scripts=[];locks=[]
    class Lock:
        def __init__(self,root):locks.append(root)
        def acquire(self,deadline):pass
        def close(self):pass
    monkeypatch.setattr(mod,'OfficeJobLock',Lock)
    monkeypatch.setattr(mod,'_container_root',lambda component:tmp_path)
    monkeypatch.setattr(mod.MacExcelSession,'_run',lambda self,body,**kw:scripts.append((body,kw)) or {})
    obj=mod.MacExcelSession.new_document(visible=False)
    assert 'make new workbook' in scripts[0][0]
    assert 'file format Excel XML file format' in scripts[0][0]
    assert obj._native.parent.parent==tmp_path and locks
    assert scripts[0][1]['bind'] is False and not obj._attached


def test_explicit_copy_timeout_reports_possible_clipboard_side_effect(tmp_path):
    import json
    obj=session(tmp_path)
    def timeout(*a,**k):raise subprocess.TimeoutExpired('osascript',60)
    obj._execute=timeout
    with pytest.raises(RuntimeError) as raised:obj.apply_structural_op({'op':'clone','target':'sheet:1/cell:A2','to':{'index':4}})
    assert raised.value.clipboard_changed is True and raised.value.clipboard_may_have_changed is True
    assert json.loads((tmp_path/'recovery.json').read_text())['clipboard_may_have_changed'] is True


def test_new_document_rebinds_native_name_after_save_as(tmp_path,monkeypatch):
    mod=module();scripts=[]
    class Lock:
        def __init__(self,root):pass
        def acquire(self,deadline):pass
        def close(self):pass
    monkeypatch.setattr(mod,'OfficeJobLock',Lock);monkeypatch.setattr(mod,'_container_root',lambda c:tmp_path)
    monkeypatch.setattr(mod.MacExcelSession,'_run',lambda self,body,**kw:scripts.append(body) or {})
    obj=mod.MacExcelSession.new_document()
    assert 'set ownedBook to workbook '+mod._quote(obj._native.name) in scripts[0]


def test_direct_write_cell_none_clears_existing_value(tmp_path):
    obj=session(tmp_path);scripts=[];obj._run=lambda body,**k:scripts.append(body) or {}
    obj.write_cell(1,1,None)
    assert 'set value of obj to ""' in scripts[0]
