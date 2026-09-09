from pathlib import Path
import hashlib,json,sys,traceback
ROOT=Path('/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description');sys.path.insert(0,str(ROOT))
from skills.WPSComposer import inspect,edit,convert_to_pdf

def main():
 source=ROOT/'docs/verification/microsoft-parity/macos-powerpoint-sessions/semantic-03-public/semantic.pptx'
 out=ROOT/'build/wps-shared-save-regression-20260909-fixed02';out.mkdir(exist_ok=False)
 report={'passed':False,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest()}
 try:
  before=inspect(source,engine='wps');(out/'before.json').write_text(json.dumps(before,ensure_ascii=False,indent=2))
  report['edit']=edit(source,engine='wps',output=out/'edited.pptx',patches=[{'target':'slide:1/shape:1','text':'WPS shared save verified'}]);assert report['edit']['ok'] and report['edit']['saved'],report['edit']
  after=inspect(out/'edited.pptx',engine='wps');(out/'after.json').write_text(json.dumps(after,ensure_ascii=False,indent=2));assert 'WPS shared save verified' in json.dumps(after)
  assert hashlib.sha256(source.read_bytes()).hexdigest()==report['source_sha256']
  report['conversion']=str(convert_to_pdf(out/'edited.pptx',out/'edited.pdf',engine='wps'))
  from pypdf import PdfReader
  reader=PdfReader(out/'edited.pdf');assert 'WPS shared save verified' in '\n'.join(p.extract_text() or '' for p in reader.pages)
  report['passed']=True;report['pages']=len(reader.pages)
 except BaseException as e:report['error']=str(e);(out/'failure.txt').write_text(traceback.format_exc())
 (out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps(report,ensure_ascii=False));return 0 if report['passed'] else 1
if __name__=='__main__':raise SystemExit(main())
