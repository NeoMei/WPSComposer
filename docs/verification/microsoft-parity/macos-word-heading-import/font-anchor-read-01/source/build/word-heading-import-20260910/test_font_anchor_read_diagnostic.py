from pathlib import Path
import importlib.util
import pytest
PATH=Path(__file__).with_name('font-anchor-read-diagnostic.py')
def module():
    assert PATH.exists(),'font anchor diagnostic absent'
    spec=importlib.util.spec_from_file_location('font_anchor_diag',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_single_scalar_anchors_use_native_utf16_bounds():
    m=module();body='😀\r'+m.SOURCE+'\r';start=3;end=len(body.encode('utf-16-le'))//2
    anchors=m.anchor_ranges(body,start,end,m.SOURCE)
    assert [a['script'] for a in anchors]==['latin','cjk','arabic','emoji']
    raw=body.encode('utf-16-le')
    for a in anchors:
        assert raw[a['start']*2:a['end']*2].decode('utf-16-le')==a['scalar']
        assert len(a['scalar'])==1
    assert anchors[-1]['end']-anchors[-1]['start']==2

@pytest.mark.parametrize('change',['surrogate','wrong_text','bool','outside'])
def test_anchor_preimage_rejects_wrong_or_invalid_bounds(change):
    m=module();body='😀\r'+m.SOURCE+'\r';start=3;end=len(body.encode('utf-16-le'))//2
    if change=='surrogate':start=1
    elif change=='wrong_text':body=body.replace('SOURCE','WRONG')
    elif change=='bool':start=True
    else:end+=1
    with pytest.raises(ValueError):m.anchor_ranges(body,start,end,m.SOURCE)

def test_read_commands_guard_exact_body_and_scalar_range():
    m=module();body=m.SOURCE+'\r'+m.IMPORTED+'\r';split=len((m.SOURCE+'\r').encode('utf-16-le'))//2
    anchors=m.all_anchors(body,0,split,split,len(body.encode('utf-16-le'))//2)
    code='\n'.join(m.font_commands(body,anchors))
    assert 'WPSC_FONT_FULL_PREIMAGE' in code and 'WPSC_FONT_ANCHOR_TEXT' in code and 'WPSC_FONT_ANCHOR_BOUNDS' in code
    assert 'set content of' not in code and 'save boundDoc' not in code
    assert all(field+' of observedFont' in code for field in m.FIELDS)

def test_empty_names_remain_observations_and_mismatches():
    m=module();style=['Arial']*5;actual=['','Arial','','','Arial']
    result=m.compare_fields(style,actual)
    assert result['equal'] is False
    assert result['differences']['ascii name']=={'style':'Arial','range':''}

def test_no_execute_guard(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()

@pytest.mark.parametrize('change',['text','surrogate_width','missing','field_type'])
def test_typed_native_observation_rejects_wrong_scalar_or_incomplete_read(change):
    m=module();body=m.SOURCE+'\r'+m.IMPORTED+'\r';split=len((m.SOURCE+'\r').encode('utf-16-le'))//2
    anchors=m.all_anchors(body,0,split,split,len(body.encode('utf-16-le'))//2)
    rows=[['styles','标题 1',m.CLONE]]
    for owner,scope,text,start,end in m.specs(anchors):
        rows.append(['font',owner,scope,text,start,end,'标题 1' if owner=='source' else m.CLONE,*(['']*5)])
    assert m.rows_valid(rows,anchors)
    if change=='text':rows[-1][3]='x'
    elif change=='surrogate_width':rows[-1][5]-=1
    elif change=='missing':rows.pop()
    else:rows[-1][7]=None
    assert not m.rows_valid(rows,anchors)
