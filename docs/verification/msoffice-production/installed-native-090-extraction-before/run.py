from pathlib import Path
import hashlib, json
import skills.WPSComposer as package
from skills.WPSComposer import generate, convert_to_pdf

def main():
    root=Path(__file__).resolve().parent
    module=Path(package.__file__).resolve()
    assert module.is_relative_to(Path('/tmp/wpscomposer-090-final/codex/plugins/wps-composer').resolve())
    out=generate(str(root/'installed-smoke.md'), format='docx', output=str(root/'installed-smoke.docx'), engine='msoffice', timeout=180)
    before=Path(out).read_bytes()
    pdf=convert_to_pdf(out, output=str(root/'installed-smoke.pdf'), engine='msoffice', timeout=180)
    assert before==Path(out).read_bytes()
    from pypdf import PdfReader
    text='\n'.join(page.extract_text() for page in PdfReader(pdf).pages)
    assert 'INSTALLED-NATIVE-END' in text
    report=dict(status='PASS',version='0.9.0',module=str(module),docx=out,pdf=pdf,source_preserved=True,pdf_end_marker=True,sha256={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (Path(out),Path(pdf))})
    (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':
    main()
