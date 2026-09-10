from pathlib import Path
import importlib.util
import xml.etree.ElementTree as E
import pytest

def m():
 p=Path(__file__).with_name('italic-face-continuation.py');s=importlib.util.spec_from_file_location('italic_cont',p);v=importlib.util.module_from_spec(s);s.loader.exec_module(v);return v

def test_actual_native_deletion_only_and_text_preservation():
 v=m();a,ax=v.package(v.BASELINE);b,bx=v.package(v.INPUT)
 assert v.italic_off_valid(a,b)
 assert v.document_signature(ax)==v.document_signature(bx)
 assert v.sha(v.INPUT)==v.INPUT_SHA and v.sha(v.BASELINE)==v.BASELINE_SHA

@pytest.mark.parametrize('change',['one_only','clone','root','default_italic','base_italic','base_chain','wrong_link','font'])
def test_bounded_deletion_rejects_other_changes(change):
 v=m();a,_=v.package(v.BASELINE);b,_=v.package(v.INPUT);before=E.fromstring(a);after=E.fromstring(b)
 index={s.get(v.W+'styleId'):s for s in after.findall(v.W+'style')}
 if change=='one_only':E.SubElement(index['1'].find(v.W+'rPr'),v.W+'i')
 elif change=='clone':index['WPSCHeadingClone'].set('unexpected','x')
 elif change=='root':after.set('unexpected','x')
 elif change=='wrong_link':index['1'].find(v.W+'link').set(v.W+'val','a0')
 elif change=='font':index['1'].find('./'+v.W+'rPr/'+v.W+'rFonts').set(v.W+'cs','Different')
 elif change=='default_italic':E.SubElement(before.find('./'+v.W+'docDefaults/'+v.W+'rPrDefault/'+v.W+'rPr'),v.W+'i')
 else:
  base=next(s for s in before.findall(v.W+'style') if s.get(v.W+'styleId')=='a')
  if change=='base_italic':E.SubElement(E.SubElement(base,v.W+'rPr'),v.W+'i')
  else:E.SubElement(base,v.W+'basedOn',{v.W+'val':'a0'})
 assert not v.italic_off_valid(E.tostring(before),E.tostring(after))

def test_continuation_does_not_repeat_off_mutation():
 source=Path(__file__).with_name('italic-face-continuation.py').read_text()
 assert "for label,value in [('restored',True)]" in source
 assert "owner.export_pdf(output/'italic-off.pdf')" in source
 assert "('off',toggle_commands(body,False))" not in source
