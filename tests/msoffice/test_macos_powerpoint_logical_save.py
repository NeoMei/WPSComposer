"""Logical SaveAs contracts over valid package publication; no native launch."""
from pathlib import Path
import shutil
import time
import zipfile
from xml.sax.saxutils import escape
import pytest
from skills.WPSComposer.scripts.msoffice import macos_powerpoint_session as module
from skills.WPSComposer.scripts.artifact_transport import snapshot_artifact_state,validate_office_package


def package(path,marker):
    p='http://schemas.openxmlformats.org/presentationml/2006/main';a='http://schemas.openxmlformats.org/drawingml/2006/main';r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
        z.writestr('_rels/.rels',f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{r}/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
        z.writestr('ppt/presentation.xml',f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>')
        z.writestr('ppt/_rels/presentation.xml.rels',f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{r}/slide" Target="slides/slide1.xml"/></Relationships>')
        z.writestr('ppt/slides/slide1.xml',f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{escape(marker)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
    validate_office_package(path,'pptx')
    return path


def read(path):
    with zipfile.ZipFile(path) as z:return z.read('ppt/slides/slide1.xml').decode()


@pytest.fixture
def bound(tmp_path,monkeypatch):
    source=package(tmp_path/'A.pptx','ORIGINAL')
    native=package(tmp_path/'native.pptx','INITIAL')
    s=module.MacPowerPointSession(source);s._deadline=time.monotonic()+120;s._path=str(native);s._entered=True
    s._source_digest=s._digest(source)
    s._logical=source;s._logical_state=snapshot_artifact_state(source,deadline=s._deadline)
    calls=[];monkeypatch.setattr(s,'_run',lambda *a,**kw:calls.append(a) or 'SAVED')
    return s,source,native,calls


def test_save_as_rebinds_only_logical_destination(bound,tmp_path):
    s,source,native,calls=bound;before=source.read_bytes();b=tmp_path/'B.pptx'
    assert s.save(b)==str(b)
    package(native,'LATER EDIT')
    assert s.save_current()==str(b)
    assert 'LATER EDIT' in read(b) and source.read_bytes()==before
    assert s._path==str(native)


def test_save_copy_preserves_logical_destination(bound,tmp_path):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';c=tmp_path/'C.pptx';s.save(b);s.save_copy(c)
    package(native,'LATER EDIT');s.save_current()
    assert 'LATER EDIT' in read(b) and 'INITIAL' in read(c) and 'ORIGINAL' in read(source)


def test_new_save_current_without_destination_fails_before_native(bound):
    s,source,native,calls=bound;s._source=None;s._source_digest=None;s._logical=None;s._logical_state=None
    with pytest.raises(ValueError,match='explicit.*destination'):s.save_current()
    assert not calls


def test_new_save_as_then_save_current_updates_published_destination(bound,tmp_path):
    s,source,native,calls=bound;s._source=None;s._source_digest=None;s._logical=None;s._logical_state=None
    b=tmp_path/'B.pptx';s.save(b);package(native,'NEW EDIT');assert s.save_current()==str(b)
    assert 'NEW EDIT' in read(b)


def test_logical_destination_conflict_fails_before_native_mutation(bound,tmp_path):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';s.save(b);calls.clear();package(b,'USER CHANGE')
    with pytest.raises((ValueError,RuntimeError),match='changed'):s.save_current()
    assert not calls and 'USER CHANGE' in read(b)


def test_logical_conflict_after_staging_preserves_changed_bytes(bound,tmp_path,monkeypatch):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';s.save(b);prior=s._logical_state
    real=s._validate_artifact
    def validate(path,fmt):
        real(path,fmt)
        if Path(path)==native:package(b,'RACING USER CHANGE')
    monkeypatch.setattr(s,'_validate_artifact',validate)
    with pytest.raises(RuntimeError,match='changed'):s.save_current()
    assert 'RACING USER CHANGE' in read(b) and native.exists() and s._logical_state==prior


def test_failed_save_as_does_not_rebind(bound,tmp_path,monkeypatch):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';c=tmp_path/'C.pptx';s.save(b);before=s._logical_state
    def fail(*a,**kw):raise RuntimeError('publication failed')
    monkeypatch.setattr(module,'publish_artifact',fail)
    with pytest.raises(RuntimeError,match='publication'):s.save(c)
    assert s._logical==b and s._logical_state==before and not c.exists()


def test_initial_save_current_rewrites_original_explicitly_and_then_repeats(bound):
    s,source,native,calls=bound
    assert s.save_current()==str(source)
    package(native,'SECOND');assert s.save_current()==str(source)
    assert 'SECOND' in read(source)

def test_conflict_during_publication_validation_preserves_destination(bound,tmp_path,monkeypatch):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';s.save(b);state=s._logical_state
    real=s._validate_artifact
    def validate(path,fmt):
        real(path,fmt)
        if Path(path)!=native:package(b,'AFTER STAGING USER CHANGE')
    monkeypatch.setattr(s,'_validate_artifact',validate)
    with pytest.raises(RuntimeError,match='changed'):s.save_current()
    assert 'AFTER STAGING USER CHANGE' in read(b) and s._logical_state==state


def test_same_bytes_replacement_is_a_conflict(bound,tmp_path):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';s.save(b);calls.clear()
    replacement=tmp_path/'swap';replacement.write_bytes(b.read_bytes());replacement.replace(b)
    with pytest.raises(ValueError,match='changed'):s.save_current()
    assert not calls


def test_save_current_after_save_as_does_not_depend_on_original_source_changes(bound,tmp_path):
    s,source,native,calls=bound;b=tmp_path/'B.pptx';s.save(b);package(source,'EXTERNAL SOURCE EDIT');package(native,'NEW B')
    assert s.save_current()==str(b)
    assert 'EXTERNAL SOURCE EDIT' in read(source) and 'NEW B' in read(b)


def test_save_current_detects_native_missing_before_false_binding(bound,tmp_path,monkeypatch):
    s,source,native,calls=bound;native.unlink();monkeypatch.setattr(s,'_validate_artifact',lambda *a:None)
    with pytest.raises((ValueError,RuntimeError,FileNotFoundError)):s.save(tmp_path/'B.pptx')
    assert s._logical==source

def test_semantic_slide_appends_bound_document_without_unrequested_view_mutation(monkeypatch):
    s=module.MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**kw:calls.append(body) or '1')
    s.add_blank_slide();code=calls[0]
    assert 'make new slide at end of ownedDoc' in code
    assert 'set slide of view' not in code and 'activate' not in code
    assert 'return slide index of currentSlide' in code


def test_owned_slide_activation_commands_compile(monkeypatch,tmp_path):
    import subprocess,sys
    if sys.platform!='darwin':pytest.skip('macOS dictionary required')
    s=module.MacPowerPointSession();calls=[];monkeypatch.setattr(s,'_run',lambda body,**kw:calls.append(body) or '1')
    s.add_blank_slide();p=tmp_path/'activate.applescript';p.write_text('tell application "Microsoft PowerPoint"\n'+calls[0]+'\nend tell')
    r=subprocess.run(['osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True)
    assert r.returncode==0,r.stderr

def test_public_edit_staging_keeps_original_transaction(bound,tmp_path,monkeypatch):
    from skills.WPSComposer.scripts import document_api
    s,source,native,calls=bound;before=source.read_bytes();out=tmp_path/'EDIT.pptx';events=[]
    monkeypatch.setattr(document_api,'open_document',lambda *a,**kw:s)
    monkeypatch.setattr(s,'preflight_save',lambda *a,**kw:None)
    monkeypatch.setattr(s,'apply_format_patch',lambda *a,**kw:(package(native,'PUBLIC EDIT') and {'accepted':['text'],'rejected':[]}))
    def close(save_changes=False):
        assert not out.exists();s._entered=False;events.append('closed')
    monkeypatch.setattr(s,'close',close)
    result=document_api.edit(source,output=out,engine='msoffice',patches=[{'target':'slide:1/shape:1','text':'PUBLIC EDIT'}])
    assert result['saved'] and source.read_bytes()==before and 'PUBLIC EDIT' in read(out)
    assert events==['closed']

def test_layout_template_keeps_actual_native_view_semantics(monkeypatch):
    from skills.WPSComposer.scripts.layout_templates import LayoutTemplate
    s=module.MacPowerPointSession();s._current_slide=999;calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**kw:calls.append(body) or 'LAYOUT')
    s.apply_layout_template(LayoutTemplate('test','content','test',[]))
    assert 'set targetSlide to slide of view of document window 1 of ownedDoc' in calls[0]
    assert 'slide 999' not in calls[0]
