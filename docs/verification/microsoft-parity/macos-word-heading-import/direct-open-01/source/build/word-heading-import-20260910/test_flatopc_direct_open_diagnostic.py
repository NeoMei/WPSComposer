from pathlib import Path
import importlib.util
import pytest

PATH=Path(__file__).with_name('flatopc-direct-open-diagnostic.py')

def module():
    assert PATH.exists(), 'direct-file diagnostic absent'
    spec=importlib.util.spec_from_file_location('direct_open',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_direct_file_is_only_open_argument_change(tmp_path):
    m=module();commands=m.commands(tmp_path/'own.xml')
    opens=[line for line in commands if 'open (POSIX file' in line]
    assert len(opens)==1
    assert 'file converter open format xmldocument serialized read only true add to recent files false confirm conversions false' in opens[0]
    assert 'file name' not in opens[0]
    assert not any('save as' in line or 'insert file' in line or 'close ' in line for line in commands)
    assert commands.index('set nativeRows to {{"open-result",resultKind,resultCount}}') < commands.index('repeat with nativeDocument in documents')

@pytest.mark.parametrize('rows,want',[
    ([['open-result','missing value',0]],True),
    ([['open-result','document',1],['document','own.xml','/private/own.xml',True,0]],True),
    ([['open-result','document',True]],False),
    ([['open-result','document',1],['document','own.xml','/private/own.xml',1,0]],False),
    ([['open-result','document',1],['document','own.xml','/private/own.xml',True]],False),
])
def test_raw_observation_accepts_zero_documents_but_not_malformed_rows(rows,want):
    assert module().observation_valid(rows) is want

def test_owned_binding_rejects_extra_or_wrong_path_documents(tmp_path):
    m=module();path=tmp_path/'own.xml'
    own=['document',path.name,str(path),True,0]
    assert m.owned_binding([['open-result','document',1],own],path)==[['binding',0,str(path),path.name]]
    assert m.owned_binding([['open-result','document',1],own,own],path) is None
    assert m.owned_binding([['open-result','document',1],['document',path.name,str(path)+'x',True,0]],path) is None
    assert m.owned_binding([['open-result','missing value',0]],path) is None

def test_no_execute_guard(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *args:pytest.fail('native entered'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
