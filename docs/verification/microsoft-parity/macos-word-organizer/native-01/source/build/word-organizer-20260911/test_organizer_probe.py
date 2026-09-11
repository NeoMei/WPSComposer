from pathlib import Path
import importlib.util
import subprocess
from zipfile import ZipFile
import xml.etree.ElementTree as ET

PROBE=Path(__file__).with_name('organizer_probe.py')
def fixture():
    spec=importlib.util.spec_from_file_location('organizer_probe',PROBE)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_donor_changes_only_one_custom_style_and_materializes_versions(tmp_path):
    m=fixture()
    for size in (17,23):
        target=tmp_path/f'donor-{size}.docx'
        m.make_input(m.SOURCE,target,size)
        with ZipFile(m.SOURCE) as a,ZipFile(target) as b:
            assert set(a.namelist())==set(b.namelist())
            assert all(a.read(n)==b.read(n) for n in a.namelist() if n!='word/styles.xml')
            original=ET.fromstring(a.read('word/styles.xml'));changed=ET.fromstring(b.read('word/styles.xml'))
            node=m.target_style(changed)
            assert node is not None and node.find('./'+m.W+'rPr/'+m.W+'sz').get(m.W+'val')==str(size*2)
            changed.remove(node)
            assert m.canonical(changed)==m.canonical(original)

def test_oracle_rejects_other_style_defaults_and_bookmark_change(tmp_path):
    m=fixture()
    with ZipFile(m.SOURCE) as z: styles=z.read('word/styles.xml')
    root=ET.fromstring(styles);root.find(m.W+'docDefaults').set('bad','1')
    assert not m.other_styles_equal(styles,ET.tostring(root))

def test_organizer_command_is_single_named_style_and_compiles(tmp_path):
    m=fixture()
    lines=m.copy_commands(Path('/tmp/owned-donor.docx'))
    script='tell application "Microsoft Word"\nset boundDoc to active document\n'+ '\n'.join(lines)+'\nend tell\n'
    p=tmp_path/'copy.applescript';p.write_text(script)
    r=subprocess.run(['osacompile','-o',str(tmp_path/'copy.scpt'),str(p)],capture_output=True,text=True)
    assert r.returncode==0,r.stderr
    assert 'organizer object styles' in script
    assert 'copy styles from template' not in script
