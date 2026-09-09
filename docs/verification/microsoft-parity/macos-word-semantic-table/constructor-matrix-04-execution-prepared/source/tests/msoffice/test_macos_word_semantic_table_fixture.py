"""Preparation-only semantic table fixture tests; no Office subprocesses."""
from __future__ import annotations

import copy
import importlib
from pathlib import Path
from types import SimpleNamespace
import xml.etree.ElementTree as ET

import pytest


def module():
    try:
        return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table')
    except ModuleNotFoundError:
        pytest.fail('Independent native semantic table fixture is missing')


def test_execute_and_fresh_output_are_required_before_any_native_work(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.safe, 'independent_inventory', lambda *a: pytest.fail('native inventory'))
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: pytest.fail('Word launch'))
    target = tmp_path/'new'
    assert m.main(['--output', str(target)]) == 2
    for value in (False, None, 1, 'yes'):
        with pytest.raises(ValueError): m.run(target, execute=value)
    assert not target.exists()
    with pytest.raises(FileExistsError): m.run(tmp_path, execute=True)


def style_rows():
    data = [['编号 ID', 'Narrative 内容', 'Flag'], ['A1', '正文 中文😀', 'V-MERGE'],
            ['A2', 'None', ''], ['H-MERGE', '', 'True']]
    result = [['style', 480.0, False, True, False,
               [135.927094400347, 220.8142163128442, 123.25868928680886]]]
    for r, row in enumerate(data, 1):
        for c, text in enumerate(row, 1):
            result.append(['cell', r, c, text+'\r\x07', 2.5, 0.0, 0.0, c-1])
    result += [['border', key, enabled, width] for key, enabled, width in
               [('top', True, 1.5), ('left', False, 0), ('bottom', True, .75),
                ('right', True, .25), ('insideHorizontal', False, 0),
                ('insideVertical', True, .25), ('headerBottom', True, .75)]]
    return result


def test_style_validator_rejects_drift_partial_duplicate_and_bool_numbers():
    m = module()
    good = style_rows()
    assert m.style_valid(good)
    for index, field, value in [(0, 1, 450), (0, 2, True), (0, 3, 1), (0, 4, True),
                                (1, 1, True), (1, 3, 'missing'), (1, 4, 0),
                                (1, 5, 1), (1, 7, 1), (16, 3, .5)]:
        bad = copy.deepcopy(good); bad[index][field] = value
        assert not m.style_valid(bad)
    assert not m.style_valid(good+good[:1])
    assert not m.style_valid(good[:-1])
    bad = copy.deepcopy(good); bad[0][5][0] = 150
    assert not m.style_valid(bad)


def native_rows():
    return [['result', '/owned.docx', 12, 180, 260, 2, 181, 181, '\r', True],
            ['outside', 'PREFIX 中文😀\r', 'SUFFIX preserved 中文😀\r', 'EXISTING TABLE 原样\r\x07', 225],
            ['table-text', '编号 ID\r\x07Narrative 内容\r\x07Flag\r\x07A1 正文 中文😀 V-MERGE A2 None H-MERGE True']]


def test_final_readback_requires_exact_scope_cr_cursor_and_preserved_outside():
    m = module()
    good = native_rows()
    assert m.result_valid(good, '/owned.docx', 12, 'EXISTING TABLE 原样\r\x07')
    for row, index, value in [(0,1,'/Owned.docx'), (0,2,13), (0,3,12), (0,4,180),
                              (0,5,True), (0,6,180), (0,7,182), (0,8,''), (0,9,1),
                              (1,1,'PREFIX changed\r'), (1,3,'EXISTING changed'), (1,4,170),
                              (2,1,'A1 A2 H-MERGE True')]:
        bad = copy.deepcopy(good); bad[row][index] = value
        assert not m.result_valid(bad, '/owned.docx', 12, 'EXISTING TABLE 原样\r\x07')
    assert not m.result_valid(good+good, '/owned.docx', 12, 'EXISTING TABLE 原样\r\x07')


def test_primitive_is_used_unmodified_and_guarded_before_middle_mutation():
    m = module()
    script = '\n'.join(m.mutation_commands('/owned.docx', 12, 19))
    assert script.index('SEMANTIC_RANGE_CHANGED') < script.index('make new table')
    assert script.index('SEMANTIC_PAGE_CHANGED') < script.index('make new table')
    assert 'create range boundDoc start 12 end 19' in script
    assert script.index('set nativeRows to {{"style"') < script.index('merge cell')
    assert script.index('row 2 column 3') < script.index('row 4 column 1')
    for forbidden in ('document 1', 'active document', 'quit', 'kill', 'clipboard', 'normal template'):
        assert forbidden not in script.lower()


@pytest.mark.parametrize('fault', ['timeout', 'bad-ack'])
def test_uncertain_observation_stops_all_followup_and_preserves_confirmed_steps(tmp_path, fault):
    m = module()
    calls = []
    class Owner:
        _bound_path = '/owned.docx'
        _closed = False
        _quarantined = False
        staging_root = tmp_path/'stage'
        def _mutation_preflight(self): calls.append('preflight')
        def _execute_structural(self, lines):
            calls.append('native')
            if fault == 'timeout': raise TimeoutError('unknown native completion')
            return [['wrong-ack']]
        def _retain(self, label): self._quarantined = True
    owner = Owner()
    report = {'status':'FAIL', 'steps':[], 'confirmed':['seed'], 'checks':{}}
    with pytest.raises((TimeoutError, AssertionError)):
        m.observe_native(owner, tmp_path, report, 'table', ['native body'], lambda rows: rows == [['ok']], mutate=True)
    first = list(calls)
    with pytest.raises(RuntimeError):
        m.observe_native(owner, tmp_path, report, 'forbidden-followup', [], lambda rows: True)
    assert calls == first == ['preflight', 'native']
    assert report['confirmed'] == ['seed'] and report['quarantined'] is True
    assert report['native_uncertainty']['owned_document']['path'] == '/owned.docx'
    assert (tmp_path/'native-uncertainty.json').is_file()


def xml_pair():
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    def node(tag, **attrs): return ET.Element(w+tag, {w+k:str(v) for k,v in attrs.items()})
    def paragraph(text, align=None):
        p=node('p')
        if align is not None:
            props=ET.SubElement(p,w+'pPr'); ET.SubElement(props,w+'ind',{w+'firstLine':'50',w+'left':'0',w+'right':'0'})
            ET.SubElement(props,w+'jc',{w+'val':align})
        r=ET.SubElement(p,w+'r'); ET.SubElement(r,w+'t').text=text
        return p
    old=node('tbl'); row=ET.SubElement(old,w+'tr'); cell=ET.SubElement(row,w+'tc'); cell.append(paragraph('EXISTING TABLE 原样'))
    before=node('document'); body=ET.SubElement(before,w+'body')
    for text in ('PREFIX 中文😀','REPLACE','SUFFIX preserved 中文😀'): body.append(paragraph(text))
    body.append(old); body.append(node('sectPr'))
    after=copy.deepcopy(before); body=after.find(w+'body'); body.remove(body[1])
    table=node('tbl'); props=ET.SubElement(table,w+'tblPr'); borders=ET.SubElement(props,w+'tblBorders')
    for key,value,width in [('top','single',12),('left','nil',None),('bottom','single',6),('right','single',2),('insideH','nil',None),('insideV','single',2)]:
        b=ET.SubElement(borders,w+key,{w+'val':value})
        if width is not None: b.set(w+'sz',str(width))
    grid=ET.SubElement(table,w+'tblGrid')
    for width in (2719,4416,2465): ET.SubElement(grid,w+'gridCol',{w+'w':str(width)})
    values=[['编号 ID','Narrative 内容','Flag'],['A1','正文 中文😀','V-MERGE'],['A2','None',''],['H-MERGE','True']]
    for ri,row_values in enumerate(values):
        row=ET.SubElement(table,w+'tr'); rp=ET.SubElement(row,w+'trPr'); ET.SubElement(rp,w+'cantSplit')
        if ri==0: ET.SubElement(rp,w+'tblHeader')
        for ci,text in enumerate(row_values):
            cell=ET.SubElement(row,w+'tc'); cp=ET.SubElement(cell,w+'tcPr')
            if ri==0:
                bs=ET.SubElement(cp,w+'tcBorders'); ET.SubElement(bs,w+'bottom',{w+'val':'single',w+'sz':'6'})
            if ci==2 and ri in (1,2): ET.SubElement(cp,w+'vMerge',({w+'val':'restart'} if ri==1 else {}))
            if ri==3 and ci==0: ET.SubElement(cp,w+'gridSpan',{w+'val':'2'})
            column=2 if ri==3 and ci==1 else ci
            cell.append(paragraph(text,['left','center','right'][column]))
    body.insert(1,table); body.insert(2,paragraph(''))
    return ET.tostring(before),ET.tostring(after)


def test_xml_gate_checks_native_grid_merge_border_indent_and_preserved_objects():
    m=module(); before,after=xml_pair()
    assert all(m.xml_checks(before,after).values())
    mutations=[(b'w="2719"',b'w="3719"','grid_widths'),
               (b'sz="12"',b'sz="4"','borders'),
               (b'firstLine="50"',b'firstLine="0"','cell_paragraphs'),
               (b'gridSpan ns0:val="2"',b'gridSpan ns0:val="3"','merge_topology'),
               (b'EXISTING TABLE',b'CHANGED TABLE','outside_preserved'),
               (b'cantSplit',b'canSplit','row_flags')]
    for old,new,key in mutations:
        assert old in after
        checks=m.xml_checks(before,after.replace(old,new,1))
        assert checks[key] is False, key


def test_pdf_gate_requires_text_order_and_real_table_width(tmp_path):
    m=module(); import fitz
    for width,expected in ((480,True),(200,False)):
        path=tmp_path/f'{width}.pdf'
        with fitz.open() as pdf:
            page=pdf.new_page(width=600,height=780)
            page.insert_text((60,50),'PREFIX')
            page.draw_line((60,80),(60+width,80),width=1.5)
            page.insert_text((60,110),'ID Narrative Flag A1 A2 V-MERGE H-MERGE None True')
            page.draw_line((60,150),(60+width,150),width=.75)
            page.insert_text((60,190),'SUFFIX preserved')
            page.insert_text((60,230),'EXISTING TABLE')
            pdf.save(path)
        checks=m.pdf_checks(path,tmp_path/f'png-{width}')
        assert checks['pdf_table_geometry'] is expected
        assert checks['pdf_text_order'] is True
        assert checks['pdf_no_images'] is True


def test_final_readback_rejects_duplicate_table_content():
    m=module(); rows=native_rows(); rows[2][1]+=' A1'
    assert not m.result_valid(rows,'/owned.docx',12,'EXISTING TABLE 原样\r\x07')


def test_xml_preservation_rejects_unexpected_outside_paragraph():
    m=module(); before,after=xml_pair(); root=ET.fromstring(after)
    body=root.find(m.W+'body'); paragraph=ET.SubElement(body,m.W+'p'); run=ET.SubElement(paragraph,m.W+'r'); ET.SubElement(run,m.W+'t').text='unexpected outside mutation'
    assert m.xml_checks(before,ET.tostring(root))['outside_preserved'] is False


def test_pdf_rejects_table_text_outside_prefix_suffix_region(tmp_path):
    m=module(); import fitz
    path=tmp_path/'outside.pdf'
    with fitz.open() as pdf:
        page=pdf.new_page(width=600,height=780)
        page.insert_text((60,50),'PREFIX'); page.draw_line((60,80),(540,80),width=1.5)
        page.insert_text((60,190),'SUFFIX preserved'); page.insert_text((60,230),'EXISTING TABLE')
        page.insert_text((60,330),'ID Narrative Flag A1 A2 V-MERGE H-MERGE None True')
        pdf.save(path)
    assert m.pdf_checks(path,tmp_path/'pages')['pdf_text_order'] is False


def test_runtime_retention_failure_preserves_primary_error_and_marks_report(tmp_path,monkeypatch):
    m=module(); stage=tmp_path/'runtime'; stage.mkdir()
    owner=SimpleNamespace(staging_root=stage,_bound_path='/owned.docx',_closed=False,_quarantined=True)
    report={'status':'FAIL','error':{'type':'TimeoutError','message':'primary native failure'}}
    def denied(*a,**k): raise OSError('local evidence copy denied')
    monkeypatch.setattr(m.shutil,'copytree',denied)
    m.retain_runtime(owner,tmp_path,report,'native-runtime')
    assert report['error']['message']=='primary native failure'
    assert report['native-runtime-retention_error']['type']=='OSError'
    assert report['native-runtime-identity']['path']=='/owned.docx'


def test_missing_dictionary_fails_locally_with_no_native_inventory(tmp_path,monkeypatch):
    m=module()
    monkeypatch.setattr(m,'SOURCES',[Path(m.__file__)])
    monkeypatch.setattr(m,'DICTIONARY',tmp_path/'absent.sdef')
    monkeypatch.setattr(m.safe,'independent_inventory',lambda *a:pytest.fail('native inventory'))
    report=m.run(tmp_path/'run',execute=True)
    assert report['status']=='FAIL' and report['error']['type']=='FileNotFoundError'
    assert (tmp_path/'run/report.json').is_file()


@pytest.mark.parametrize('kind', ['empty', 'pageBreakBefore', 'sectPr', 'run-break'])
def test_xml_rejects_extra_textless_structural_paragraph_outside_splice(kind):
    m=module(); before,after=xml_pair(); root=ET.fromstring(after); body=root.find(m.W+'body')
    paragraph=ET.Element(m.W+'p')
    if kind=='run-break':
        run=ET.SubElement(paragraph,m.W+'r'); ET.SubElement(run,m.W+'br',{m.W+'type':'page'})
    elif kind!='empty':
        props=ET.SubElement(paragraph,m.W+'pPr'); ET.SubElement(props,m.W+kind)
    body.insert(len(body)-1,paragraph)
    assert m.xml_checks(before,ET.tostring(root))['outside_preserved'] is False


@pytest.mark.parametrize('mutation', ['remove', 'pageBreakBefore', 'sectPr'])
def test_xml_preserves_preexisting_empty_paragraph_structure(mutation):
    m=module(); before,after=xml_pair(); old,new=ET.fromstring(before),ET.fromstring(after)
    for root in (old,new):
        body=root.find(m.W+'body'); paragraph=ET.Element(m.W+'p')
        props=ET.SubElement(paragraph,m.W+'pPr'); ET.SubElement(props,m.W+'spacing',{m.W+'after':'120'})
        body.insert(len(body)-1,paragraph)
    before,after=ET.tostring(old),ET.tostring(new)
    assert m.xml_checks(before,after)['outside_preserved'] is True
    body=new.find(m.W+'body'); blank=body[-2]
    if mutation=='remove': body.remove(blank)
    else: ET.SubElement(blank.find(m.W+'pPr'),m.W+mutation)
    assert m.xml_checks(before,ET.tostring(new))['outside_preserved'] is False


@pytest.mark.parametrize('mutation', ['remove', 'move', 'duplicate', 'pageBreakBefore', 'sectPr', 'run-break', 'spacing'])
def test_only_one_exact_table_tail_paragraph_delta_is_allowed(mutation):
    m=module(); before,after=xml_pair(); root=ET.fromstring(after); body=root.find(m.W+'body'); tail=body[2]
    if mutation=='remove': body.remove(tail)
    elif mutation=='move': body.remove(tail); body.insert(len(body)-1,tail)
    elif mutation=='duplicate': body.insert(3,copy.deepcopy(tail))
    elif mutation=='run-break':
        run=ET.SubElement(tail,m.W+'r'); ET.SubElement(run,m.W+'br',{m.W+'type':'page'})
    else:
        props=ET.SubElement(tail,m.W+'pPr'); ET.SubElement(props,m.W+mutation)
    assert m.xml_checks(before,ET.tostring(root))['outside_preserved'] is False


def test_native_count_preconditions_are_grouped_before_comparison():
    script='\n'.join(module().mutation_commands('/owned.docx',12,19))
    assert 'if (count sections of boundDoc) is not 1 then' in script
    assert 'if (count tables of boundDoc) is not 1 then' in script
    assert 'if count sections' not in script and 'if count tables' not in script
