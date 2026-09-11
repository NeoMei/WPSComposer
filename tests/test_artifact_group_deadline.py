import time
import pytest
from skills.WPSComposer.scripts import artifact_transport as transport


def test_group_expired_deadline_has_no_publication(tmp_path):
    source = tmp_path/'source'; source.write_bytes(b'new')
    target = tmp_path/'target'; target.write_bytes(b'old')
    with pytest.raises(TimeoutError):
        transport.publish_artifact_group([(source,target,True,lambda path: None)], deadline=time.monotonic()-1)
    assert target.read_bytes() == b'old'
    assert sorted(p.name for p in tmp_path.iterdir()) == ['source','target']


def test_group_expiry_during_validation_cleans_prepared_file(tmp_path, monkeypatch):
    source = tmp_path/'source'; source.write_bytes(b'new')
    target = tmp_path/'target'; target.write_bytes(b'old')
    clock = [10.0]
    monkeypatch.setattr(transport.time,'monotonic',lambda: clock[0])
    def validator(path):
        if path != source: clock[0] = 30
    with pytest.raises(TimeoutError):
        transport.publish_artifact_group([(source,target,True,validator)], deadline=20)
    assert target.read_bytes() == b'old'
    assert sorted(p.name for p in tmp_path.iterdir()) == ['source','target']


def test_group_expiry_after_first_publication_restores_entire_set(tmp_path, monkeypatch):
    sources = [tmp_path/'source1', tmp_path/'source2']
    targets = [tmp_path/'target1', tmp_path/'target2']
    for path in sources: path.write_bytes(b'new')
    for path in targets: path.write_bytes(b'old')
    clock = [10.0]
    monkeypatch.setattr(transport.time, 'monotonic', lambda: clock[0])
    replace = transport.os.replace
    def publish_then_expire(source, target):
        replace(source, target)
        if target == targets[0]: clock[0] = 30.0
    monkeypatch.setattr(transport.os, 'replace', publish_then_expire)
    entries = [(source, target, True, lambda p: None) for source,target in zip(sources,targets)]
    with pytest.raises(TimeoutError):
        transport.publish_artifact_group(entries, deadline=20)
    assert [path.read_bytes() for path in targets] == [b'old', b'old']
    assert sorted(tmp_path.iterdir()) == sorted(sources + targets)
