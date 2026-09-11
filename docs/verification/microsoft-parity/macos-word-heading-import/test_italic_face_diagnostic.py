from pathlib import Path
import importlib.util
from zipfile import ZipFile
import xml.etree.ElementTree as E
import pytest

def m():
 s=importlib.util.spec_from_file_location('italic_probe',Path(__file__).with_name('italic-face-diagnostic.py'));v=importlib.util.module_from_spec(s);s.loader.exec_module(v);return v

def baseline(v):
 with ZipFile(v.INPUT) as z:return z.read('word/styles.xml')

def changed(v):
 root=E.fromstring(baseline(v))
 for identifier in ('1','10'):
  style=next(s for s in root.findall(v.W+'style') if s.get(v.W+'styleId')==identifier)
  style.find('./'+v.W+'rPr/'+v.W+'i').set(v.W+'val','0')
 return root

def test_only_reciprocal_pair_italic_change():
 v=m();assert v.italic_off_valid(baseline(v),E.tostring(changed(v)))

@pytest.mark.parametrize('mutation',['clone','defaults','link','hidden','other_property','one_style'])
def test_other_changes_rejected(mutation):
 v=m();root=changed(v);styles=root.findall(v.W+'style');source=next(s for s in styles if s.get(v.W+'styleId')=='1')
 if mutation=='clone':styles[-1].set('unexpected','yes')
 elif mutation=='defaults':root.set('unexpected','yes')
 elif mutation=='link':source.find(v.W+'link').set(v.W+'val','a0')
 elif mutation=='hidden':E.SubElement(source.find('./'+v.W+'rPr/'+v.W+'i'),v.W+'t').text='x'
 elif mutation=='other_property':source.find('./'+v.W+'rPr/'+v.W+'sz').set(v.W+'val','40')
 else:next(s for s in styles if s.get(v.W+'styleId')=='10').find('./'+v.W+'rPr/'+v.W+'i').attrib.clear()
 assert not v.italic_off_valid(baseline(v),E.tostring(root))

def test_commands_change_only_source_italic_and_guard_preimage():
 v=m();cmd='\n'.join(v.toggle_commands('EXACT😀\r',False));assert 'set italic of font object of sourceStyle to false' in cmd
 assert 'set italic of font object of linkedStyle' not in cmd and 'EXACT😀' in cmd
 assert 'WPSC_ITALIC_PREIMAGE' in cmd and 'WPSC_ITALIC_PAIR_PREIMAGE' in cmd

def test_raw_body_signature_preserves_all_paragraphs():
 v=m()
 with ZipFile(v.INPUT) as z:xml=z.read('word/document.xml')
 sig=v.document_signature(xml);root=E.fromstring(xml);root.find(v.W+'body').findall(v.W+'p')[1].find('.//'+v.W+'t').text='changed'
 assert v.document_signature(E.tostring(root))!=sig

def test_typed_native_state_rejects_numeric_boolean():
 v=m();row=['state',True,'body',4,0,0,True,True,True]
 assert v.valid_state([row],'body',True)
 row[6]=1;assert not v.valid_state([row],'body',True)

def test_signature_rejects_fields_and_preserves_direct_fonts():
 v=m()
 with ZipFile(v.INPUT) as z:xml=z.read('word/document.xml')
 root=E.fromstring(xml);run=root.find(v.W+'body').findall(v.W+'p')[1].find(v.W+'r')
 rpr=run.find(v.W+'rPr')
 if rpr is None:rpr=E.SubElement(run,v.W+'rPr')
 E.SubElement(rpr,v.W+'rFonts',{v.W+'cs':'Changed'})
 assert v.document_signature(E.tostring(root))!=v.document_signature(xml)
 E.SubElement(run,v.W+'fldChar')
 with pytest.raises(ValueError):v.document_signature(E.tostring(root))
