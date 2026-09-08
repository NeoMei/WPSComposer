# coding: utf-8
from pathlib import Path
import json,zipfile,xml.etree.ElementTree as E
from skills.WPSComposer import generate,convert_to_pdf
import pdfplumber
root=Path(__file__).resolve().parent
reports={}
for scheme in ['chinese-formal','hybrid-bid']:
 md=root/(scheme+'.md');md.write_text('---\ntitle: Native outline schemes\ntoc: true\nheading_numbering: '+scheme+'\n---\n# 第一章 ChapterOne\n\n## 第一节 ChildOne\n\n### 一、GrandOne\n\n#### （一）DeepOne\n\n# 第二章 ChapterTwo\n\n## 第一节 ChildTwo\n\n### 一、GrandTwo\n\n#### （一）DeepTwo\n',encoding='utf-8')
 if scheme == 'hybrid-bid':
  text=md.read_text().replace('第一节 ChildOne','1.1 ChildOne').replace('一、GrandOne','1.1.1 GrandOne').replace('（一）DeepOne','关键工法001：DeepOne').replace('第一节 ChildTwo','2.1 ChildTwo').replace('一、GrandTwo','2.1.1 GrandTwo').replace('（一）DeepTwo','关键工法001：DeepTwo');md.write_text(text,encoding='utf-8')
 d=generate(str(md),format='docx',output=str(root/(scheme+'.docx')),engine='wps',timeout=600)
 p=convert_to_pdf(d,output=str(root/(scheme+'.pdf')),engine='wps',timeout=300)
 with pdfplumber.open(p) as pdf:text='\n'.join(page.extract_text() or '' for page in pdf.pages)
 reports[scheme]={'docx':d,'pdf':p,'visible_text':text};print(scheme,text,flush=True)
(root/'report.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
