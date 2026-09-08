import sys,json,time
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer import generate
from fixtures.verify_msoffice_production import fixture_content, inspect_pdf
root=Path('build/msoffice-production/public-pdf-01').resolve()
root.mkdir(exist_ok=False)
content,spec=fixture_content('representative')
source=root/'source.md';source.write_text(content,encoding='utf-8')
pdf=root/'generated.pdf'
report={'status':'RUNNING','public_api':'generate(format=pdf, engine=msoffice)','timeout':600}
try:
    started=time.monotonic()
    result=generate(str(source),format='pdf',output=str(pdf),engine='msoffice',timeout=600,open_result=False)
    report.update(returned_requested_path=Path(result).resolve()==pdf,pdf=inspect_pdf(pdf,spec),no_unrequested_docx=not list(root.glob('*.docx')),elapsed_seconds=round(time.monotonic()-started,3))
    report['status']='PASS' if report['returned_requested_path'] and report['no_unrequested_docx'] and all(report['pdf']['checks'].values()) else 'FAIL'
except Exception as exc:
    import traceback
    report.update(status='FAIL',error=repr(exc),traceback=traceback.format_exc())
finally:
    (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
raise SystemExit(0 if report['status']=='PASS' else 1)
