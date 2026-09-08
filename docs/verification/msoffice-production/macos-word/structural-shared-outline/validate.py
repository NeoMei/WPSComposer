# coding: utf-8
from pathlib import Path
import hashlib,json,re,unicodedata,zipfile,xml.etree.ElementTree as E
import fitz
root=Path(__file__).resolve().parent
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
def normalize(text): return re.sub(r'\s+','',unicodedata.normalize('NFKC',text))
checks={}
expected={'before': [('一级章节验收',1,1)],'inserted':[('新增原生编号章节',1,1),('一级章节验收',2,2)],'reopened':[('新增原生编号章节',1,1),('一级章节验收',2,2)],'deleted':[('一级章节验收',1,1)],'deleted-reopened':[('一级章节验收',1,1)]}
for phase,entries in expected.items():
 log=(root/(phase+'.log')).read_text(); checks[phase+'_closed_and_preserved']='WPSC_CLEAN' in log and 'WPSC_OK\t0' in log
 with zipfile.ZipFile(root/(phase+'.docx')) as z: tree=E.fromstring(z.read('word/document.xml'))
 paragraphs=[''.join(t.text or '' for t in p.findall('.//w:t',ns)) for p in tree.findall('.//w:p',ns)]
 for label,number,page in entries:
  checks[phase+'_'+label+'_native_readback']=f'{label}\t{number}\t{number}\t{page}' in log
  checks[phase+'_'+label+'_cached_toc']=f'{number}{label}{page}' in paragraphs
 if phase not in ('inserted','reopened'): checks[phase+'_removed_heading_absent']=not any('新增原生编号章节' in p for p in paragraphs)
 if phase in ('before','inserted','deleted-reopened'):
  doc=fitz.open(root/(phase+'.pdf')); toc=normalize(doc[1].get_text())
  for label,number,page in entries:
   checks[phase+'_'+label+'_visible_toc']=bool(re.search(str(number)+re.escape(label)+r'[.·…]*'+str(page),toc))
   checks[phase+'_'+label+'_actual_body_page']=str(number)+label in normalize(doc[page+1].get_text())
   footer=normalize(doc[page+1].get_text(clip=fitz.Rect(0,doc[page+1].rect.height-80,doc[page+1].rect.width,doc[page+1].rect.height)))
   checks[phase+'_'+label+'_actual_footer']=footer==str(page)
for phase in expected:
 log=(root/(phase+'.log')).read_text()
 chapter='2' if phase in ('inserted','reopened') else '1'
 page=2 if chapter=='2' else 1
 for depth,label in [(2,'二级章节验收'),(3,'三级章节验收'),(4,'四级章节验收')]:
  number=chapter+'.1'*(depth-1)
  checks[phase+'_'+label+'_native_descendant']=f'{label}\t{number}\t{page}' in log
  if phase in ('before','inserted','deleted-reopened'):
   doc=fitz.open(root/(phase+'.pdf'))
   checks[phase+'_'+label+'_visible_descendant']=number+label in normalize(doc[page+1].get_text())
   if depth<=3:
    checks[phase+'_'+label+'_visible_toc']=bool(re.search(re.escape(number+label)+r'[.·…]*'+str(page),normalize(doc[1].get_text())))
report=json.loads((root/'report.json').read_text());report.update(status='PASS' if all(checks.values()) else 'FAIL',checks=checks,files={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in root.iterdir() if p.is_file() and p.name not in ('report.json','validation.log')},production_refresh_sha256=hashlib.sha256(Path('skills/WPSComposer/scripts/msoffice/macos_script.py').read_bytes()).hexdigest())
(root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
print(json.dumps(report,ensure_ascii=False,indent=2));assert all(checks.values())
