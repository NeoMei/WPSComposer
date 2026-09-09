"""Owned Word logical-save contract over valid packages, no Office process."""
import time
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile
import pytest
from skills.WPSComposer.scripts.msoffice import macos_word_session as module
from skills.WPSComposer.scripts.artifact_transport import ArtifactTransportError,snapshot_artifact_state,validate_office_package


def package(path,marker):
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>'+escape(marker)+'</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')
    validate_office_package(path,'docx');return path


def read(path):
    with zipfile.ZipFile(path) as z:return z.read('word/document.xml').decode()


@pytest.fixture
def bound(tmp_path,monkeypatch):
    source=package(tmp_path/'A.docx','ORIGINAL');native=package(tmp_path/'native.docx','INITIAL')
    s=module.MacWordSession();s._owns_doc=True;s._bound_path=str(native);s._source_path=source;s.staging_root=tmp_path;s._deadline=time.monotonic()+120;s._source_digest=s._digest(source)
    s._logical_path=source;s._logical_state=snapshot_artifact_state(source,deadline=s._deadline)
    calls=[];monkeypatch.setattr(s,'_execute',lambda *a,**kw:calls.append(a) or [['ok']]);return s,source,native,calls


def test_save_as_then_current_updates_only_new_logical_target(bound,tmp_path):
    s,a,n,calls=bound;before=a.read_bytes();b=tmp_path/'B.docx';s.save(b);package(n,'LATER');assert s.save_current()==str(b)
    assert 'LATER' in read(b) and a.read_bytes()==before and s._bound_path==str(n)


def test_save_copy_preserves_previous_logical_target(bound,tmp_path):
    s,a,n,calls=bound;b=tmp_path/'B.docx';c=tmp_path/'C.docx';s.save(b);s.save_copy(c);package(n,'LATER');s.save_current()
    assert 'LATER' in read(b) and 'INITIAL' in read(c) and 'ORIGINAL' in read(a)


def test_new_document_requires_explicit_target_then_updates_it(bound,tmp_path):
    s,a,n,calls=bound;s._source_path=None;s._source_digest=None;s._logical_path=None;s._logical_state=None
    with pytest.raises(ValueError,match='explicit'):s.save_current()
    assert not calls
    b=tmp_path/'B.docx';s.save(b);package(n,'NEW LATER');assert s.save_current()==str(b)
    assert 'NEW LATER' in read(b)


def test_destination_change_is_rejected_before_native_save(bound,tmp_path):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);package(b,'USER CHANGE');calls.clear()
    with pytest.raises((ValueError,RuntimeError),match='changed'):s.save_current()
    assert not calls and 'USER CHANGE' in read(b)


def test_same_byte_inode_replacement_is_rejected_before_native_save(bound,tmp_path):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);swap=tmp_path/'swap';swap.write_bytes(b.read_bytes());swap.replace(b);calls.clear()
    with pytest.raises((ValueError,RuntimeError),match='changed'):s.save_current()
    assert not calls


def test_after_staging_race_preserves_changed_destination_and_binding(bound,tmp_path,monkeypatch):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);prior=s._logical_state
    real=module.validate_before_deadline
    def validate(spec,path,deadline):
        real(spec,path,deadline)
        if Path(path)!=n:package(b,'RACING USER CHANGE')
    monkeypatch.setattr(module,'validate_before_deadline',validate)
    with pytest.raises(RuntimeError,match='changed'):s.save_current()
    assert 'RACING USER CHANGE' in read(b) and s._logical_state==prior and s._retain_evidence


def test_failed_save_as_does_not_rebind(bound,tmp_path,monkeypatch):
    s,a,n,calls=bound;b=tmp_path/'B.docx';c=tmp_path/'C.docx';s.save(b);state=s._logical_state
    monkeypatch.setattr(module,'publish_artifact',lambda *a,**kw:(_ for _ in ()).throw(RuntimeError('publication failed')))
    with pytest.raises(RuntimeError,match='publication'):s.save(c)
    assert s._logical_path==b and s._logical_state==state and s._retain_evidence


def test_initial_save_current_can_repeat_and_explicitly_updates_original(bound):
    s,a,n,calls=bound;assert s.save_current()==str(a);package(n,'SECOND');s.save_current();assert 'SECOND' in read(a)


def test_original_source_changes_do_not_retarget_logical_save(bound,tmp_path):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);package(a,'EXTERNAL SOURCE CHANGE');package(n,'NEW B')
    assert s.save_current()==str(b) and 'EXTERNAL SOURCE CHANGE' in read(a) and 'NEW B' in read(b)


def test_preflight_conflict_retains_unsaved_private_recovery(bound,tmp_path):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);package(n,'UNSAVED RECOVERY');package(b,'EXTERNAL');calls.clear()
    with pytest.raises(ValueError,match='changed'):s.preflight_save()
    assert not calls and s._retain_evidence and 'UNSAVED RECOVERY' in read(n)


def test_failed_save_copy_keeps_previous_binding(bound,tmp_path,monkeypatch):
    s,a,n,calls=bound;b=tmp_path/'B.docx';s.save(b);state=s._logical_state
    monkeypatch.setattr(module,'publish_artifact',lambda *a,**kw:(_ for _ in ()).throw(RuntimeError('copy failed')))
    with pytest.raises(RuntimeError,match='copy'):s.save_copy(tmp_path/'C.docx')
    assert s._logical_path==b and s._logical_state==state and s._retain_evidence


def test_unstable_logical_fingerprint_retains_private_recovery(tmp_path, monkeypatch):
    logical = package(tmp_path / "logical.docx", "published")
    stage = tmp_path / "stage"
    stage.mkdir()
    private = package(stage / "private.docx", "UNPUBLISHED EDIT")
    session = module.MacWordSession()
    session._owns_doc = True
    session._bound_path = str(private)
    session.staging_root = stage
    session._logical_path = logical
    session._logical_state = snapshot_artifact_state(logical)
    session._deadline = time.monotonic() + 30
    session._execute = lambda *_args, **_kwargs: [["ok"]]

    def unstable(*_args, **_kwargs):
        raise ArtifactTransportError(
            "ARTIFACT_STATE_UNSTABLE", "logical bytes changed during fingerprint"
        )

    monkeypatch.setattr(module, "snapshot_artifact_state", unstable)

    try:
        with session:
            session.save_current()
    except ArtifactTransportError as error:
        assert error.code == "ARTIFACT_STATE_UNSTABLE"
    else:
        raise AssertionError("unstable destination fingerprint was accepted")

    assert private.exists(), "context close deleted the only private unpublished copy"
    assert stage.exists()
    assert session._retain_evidence is True



def test_missing_destination_close_can_retry_save_then_release_lock(bound,tmp_path):
    class Lock:
        closed = False
        def close(self): self.closed = True
    s,a,n,calls=bound;s._logical_path=None;s._logical_state=None;s.lock=Lock()
    stage=tmp_path/'stage';stage.mkdir();private=stage/'native.docx';n.replace(private);n=private
    s.staging_root=stage;s._bound_path=str(private)
    with pytest.raises(ValueError,match='explicit'):s.close(save_changes=True)
    assert not s._closed and not s.lock.closed and n.exists()
    b=tmp_path/'B.docx';s.save(b);s.close()
    assert s._closed and s.lock.closed and b.exists() and not stage.exists()
