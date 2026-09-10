from pathlib import Path
import importlib.util
from copy import deepcopy
import xml.etree.ElementTree as ET
import pytest
PATH=Path(__file__).with_name('native-donor-heading-import-v3.py')
def module():
    assert PATH.exists(),'v3 control acceptance absent'
    spec=importlib.util.spec_from_file_location('donor_v3',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def control_document(m):
    root=ET.Element(m.W+'document');body=ET.SubElement(root,m.W+'body')
    for index in range(1,5):
        p=ET.SubElement(body,m.W+'p')
        if index in (1,4):
            ET.SubElement(ET.SubElement(p,m.W+'pPr'),m.W+'pStyle',{m.W+'val':'H1' if index==1 else 'CLONE'})
            run=ET.SubElement(p,m.W+'r');ET.SubElement(ET.SubElement(run,m.W+'rPr'),m.W+'rFonts',{m.W+'ascii':'Apple Color Emoji'})
            ET.SubElement(run,m.W+'t').text=m.FRESH
    return root

def test_identical_fallback_is_required_without_font_whitelist():
    m=module();root=control_document(m);xml=ET.tostring(root)
    assert m.control_signatures(xml,'H1','CLONE')[0]==m.control_signatures(xml,'H1','CLONE')[1]
    clone=root.find(m.W+'body').findall(m.W+'p')[3]
    clone.find('.//'+m.W+'rFonts').set(m.W+'ascii','Unknown Different Font')
    left,right=m.control_signatures(ET.tostring(root),'H1','CLONE');assert left!=right

@pytest.mark.parametrize('change',['wrong_style','extra_ppr','dropped_szCs','wrong_text'])
def test_direct_format_and_identity_difference_cannot_escape(change):
    m=module();root=control_document(m);clone=root.find(m.W+'body').findall(m.W+'p')[3]
    if change=='wrong_style':clone.find('.//'+m.W+'pStyle').set(m.W+'val','H1')
    elif change=='extra_ppr':ET.SubElement(clone.find(m.W+'pPr'),m.W+'ind',{m.W+'left':'5'})
    elif change=='dropped_szCs':ET.SubElement(root.find('.//'+m.W+'rPr'),m.W+'szCs',{m.W+'val':'48'})
    else:clone.find('.//'+m.W+'t').text='WRONG'
    try:left,right=m.control_signatures(ET.tostring(root),'H1','CLONE')
    except ValueError:return
    assert left!=right

def test_two_controls_require_indexed_empty_carriers():
    m=module();root=ET.Element(m.W+'document');body=ET.SubElement(root,m.W+'body')
    for _ in range(4):ET.SubElement(body,m.W+'p')
    assert m.carriers_empty(ET.tostring(root),(1,4))
    ET.SubElement(body[0],m.W+'pPr');assert not m.carriers_empty(ET.tostring(root),(1,4))

def test_constructor_is_same_for_both_styles_no_resets_or_cr_insertion():
    m=module();body='\rREPLACE\rSUFFIX\r\r'
    for index,start,style in [(1,0,'source'),(4,len(body.encode('utf-16-le'))//2-1,'clone')]:
        code='\n'.join(m.control_constructor(body,index,start,style))
        assert code.index('set style of carrierRange')<code.index('set content of freshInsert')
        assert 'reset' not in code and 'WPSC_V3_EMPTY_CARRIER' in code
        assert next(x for x in code.splitlines() if x.startswith('set content of freshInsert'))=='set content of freshInsert to '+m.apple_string(m.FRESH)

def test_duplicate_text_anchor_ranges_are_exact_indices_and_surrogates():
    m=module();body=m.FRESH+'\r'+m.IMPORTED+'\rSUFFIX\r'+m.FRESH+'\r'
    anchors=m.control_anchors(body)
    assert anchors['source'][0]['start']==0
    assert anchors['clone'][0]['start']>anchors['source'][0]['end']
    for owner in anchors:assert anchors[owner][-1]['end']-anchors[owner][-1]['start']==2

def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()

def test_size_and_spacing_challenge_requires_control_change_and_clone_stability():
    m=module();before=[['challenge',17,24,17,24,17,24,17,24]]
    after=[['challenge',28,31,28,31,17,24,17,24]]
    assert m.challenge_changed(before,after)
    assert not m.challenge_changed(before,before)
    after[0][7]=28;assert not m.challenge_changed(before,after)
    before[0][1]=float('inf');assert not m.challenge_rows_valid(before)

def test_source_challenge_xml_rejects_any_unrelated_style_change():
    m=module();before=m.SOURCE_STYLES.read_bytes();root=ET.fromstring(before)
    source=m.paragraph_style_by_name(root.findall(m.W+'style'),'heading 1')
    size=source.find('./'+m.W+'rPr/'+m.W+'sz');size.set(m.W+'val',str(int(size.get(m.W+'val'))+22))
    space=source.find('./'+m.W+'pPr/'+m.W+'spacing');space.set(m.W+'before',str(int(space.get(m.W+'before'))+140))
    assert m.source_only_challenge_xml(before,ET.tostring(root))
    normal=m.paragraph_style_by_name(root.findall(m.W+'style'),'Normal');normal.set('changed','yes')
    assert not m.source_only_challenge_xml(before,ET.tostring(root))

def test_nonempty_anchor_equality_rejects_unknown_family_change_or_empty_name():
    m=module();body=m.FRESH+'\r'+m.IMPORTED+'\rSUFFIX\r'+m.FRESH+'\r';anchors=m.control_anchors(body)
    rows=[['styles','标题 1',m.CLONE]]
    for owner,scope,text,start,end in m.font.specs(anchors):
        rows.append(['font',owner,scope,text,start,end,'标题 1' if owner=='source' else m.CLONE,*(['Exact Family']*5)])
    assert m.anchors_equal(rows,anchors)
    rows[-1][7]='Unknown Different Font';assert not m.anchors_equal(rows,anchors)
    rows[-1][7]='';assert not m.anchors_nonempty(rows,anchors)

def test_clone_font_stability_checks_all_retained_values():
    m=module();fonts=[[label,['number',17],['number',17],True] for label in m.heading.expected_labels('font-properties')]
    before={'fonts':fonts,'anchors':[['font','clone','emoji','😀',1,3,m.CLONE,'Arial']],'tabs':[['clone',1,40,'right','none',False]]}
    after=deepcopy(before);assert m.clone_observations_stable(before,after)
    after['fonts'][2][2]=['number',18];after['fonts'][2][3]=False
    assert not m.clone_observations_stable(before,after)

@pytest.mark.parametrize('change',['defaults','root_attribute','nonstyle_child'])
def test_source_challenge_rejects_nonstyle_xml_drift(change):
    m=module();before=m.SOURCE_STYLES.read_bytes();root=ET.fromstring(before)
    source=m.paragraph_style_by_name(root.findall(m.W+'style'),'heading 1')
    size=source.find('./'+m.W+'rPr/'+m.W+'sz');size.set(m.W+'val',str(int(size.get(m.W+'val'))+22))
    space=source.find('./'+m.W+'pPr/'+m.W+'spacing');space.set(m.W+'before',str(int(space.get(m.W+'before'))+140))
    if change=='defaults':root.find('.//'+m.W+'docDefaults//'+m.W+'szCs').set(m.W+'val','55')
    elif change=='root_attribute':root.set('changed','yes')
    else:ET.SubElement(root,m.W+'latentStyles',{'changed':'yes'})
    assert not m.source_only_challenge_xml(before,ET.tostring(root))

def test_v3_requires_exactly_nineteen_effective_tabs():
    m=module()
    def rows(n):return [['tabs','paragraph','标题 1',m.CLONE,n,n]]+[[side,i,i*36,'align tab left','tab leader spaces',False] for side in ('source','clone') for i in range(1,n+1)]
    assert m.fixture_tabs_valid(rows(19))
    for n in (0,18,20):assert not m.fixture_tabs_valid(rows(n))
    wrong=rows(19);wrong[-1][2]+=1;assert not m.fixture_tabs_valid(wrong)

@pytest.mark.parametrize('change',['defaults','root_attribute','nonstyle_child'])
def test_restoration_requires_full_baseline_styles_tree(change):
    m=module();before=m.SOURCE_STYLES.read_bytes();root=ET.fromstring(before)
    if change=='defaults':root.find('.//'+m.W+'docDefaults//'+m.W+'szCs').set(m.W+'val','99')
    elif change=='root_attribute':root.set('changed','yes')
    else:ET.SubElement(root,m.W+'latentStyles',{'changed':'yes'})
    assert m.complete_styles_equal(before,before)
    assert not m.complete_styles_equal(before,ET.tostring(root))
