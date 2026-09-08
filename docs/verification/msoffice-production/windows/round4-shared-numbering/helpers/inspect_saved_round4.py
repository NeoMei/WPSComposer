from pathlib import Path
import sys,json,re,zipfile,xml.etree.ElementTree as ET
import fitz
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from fixtures.verify_msoffice_production import NS,W,HEADINGS,fixture_content,inspect_docx
root=Path('build/msoffice-production/representative-04')
_,spec=fixture_content('representative')
with zipfile.ZipFile(root/'generated.docx') as package:
    xml=ET.fromstring(package.read('word/document.xml'))
fields=[n.text or '' for n in xml.findall('.//w:instrText',NS)]
headers=xml.findall('.//w:tbl/w:tr/w:trPr/w:tblHeader',NS)
checks={'docx_native_contract':all(inspect_docx(root/'generated.docx',spec)['checks'].values()),'header_on_off_value_enabled':bool(headers) and all(n.get(W+'val','true').lower() not in ['0','false','off'] for n in headers),'toc_field_uses_three_outline_levels':any(re.search(r'TOC.*\\o\s+"1-3"',f) for f in fields),'toc_has_three_native_page_references':sum('PAGEREF' in f for f in fields)==3}
with fitz.open(root/'converted.pdf') as pdf:
    pages=[p.get_text() for p in pdf]; compact=[re.sub(r'\s+','',t) for t in pages]
    spans=[s for b in pdf[2].get_text('dict')['blocks'] for line in b.get('lines',[]) for s in line['spans']]
    first=next(s for s in spans if s['text'].startswith('这是原生排版正文'))
    second=next(s for s in spans if s['text'].startswith('中文内容完整保留'))
    indent=first['origin'][0]-second['origin'][0]
    checks['pdf_body_fangsong_12pt_and_24pt_indent']=all(s['font']=='FangSong' and abs(s['size']-12)<.01 for s in [first,second]) and abs(indent-24)<.15
    sizes=[next(s['size'] for s in spans if s['text'].strip()==h) for h in HEADINGS]
    checks['six_pdf_heading_sizes_with_word_export_rounding']=all(abs(actual-expected)<.15 for actual,expected in zip(sizes,[16,15,15,14,14,12]))
    checks['four_numbered_headings_on_body_page']=all(number+heading in compact[2] for number,heading in zip(['1','1.1','1.1.1','1.1.1.1'],HEADINGS))
    checks['toc_page_numbers_match_logical_body_page_one']=all(re.search(re.escape(h)+r'\.+1',compact[1]) and h in compact[2] for h in HEADINGS[:3])
    checks['all_36_table_records_once']=all(''.join(compact).count(f'长表第{i:02d}项')==1 for i in range(1,37))
    checks['table_header_on_all_table_pages']=all(all(h in compact[i] for h in ['序号','验收项目','结果']) for i in range(2,5))
    checks['text_inside_page_bounds']=all(s['bbox'][0]>=-1 and s['bbox'][1]>=-1 and s['bbox'][2]<=p.rect.width+1 and s['bbox'][3]<=p.rect.height+1 for p in pdf for b in p.get_text('dict')['blocks'] for line in b.get('lines',[]) for s in line['spans'])
report={'checks':checks,'all_passed':all(checks.values()),'body_indent_pt':indent,'pdf_heading_sizes_pt':sizes,'pdf_heading_size_tolerance_pt':.15,'rendered_pages':5,'visual_review':'PENDING separate contact sheet inspection'}
(root/'independent-validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
raise SystemExit(0 if report['all_passed'] else 1)
