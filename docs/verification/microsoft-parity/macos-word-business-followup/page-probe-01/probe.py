from pathlib import Path
import json, shutil
from skills.WPSComposer import create_document

def main():
    out=Path(__file__).parent
    with create_document('writer',engine='msoffice',visible=False) as s:
        s.add_paragraph('Unchanged body paragraph')
        rows=s._execute(['set pagePart to get footer (section 1 of boundDoc) index header footer primary',
            'set content of text object of pagePart to "Page "',
            'make new page number at pagePart with properties {alignment:align page number left}',
            'set nativeRows to {{content of text object of boundDoc,content of text object of pagePart}}'])
        (out/'readback.json').write_text(json.dumps(rows,ensure_ascii=False))
        s.save_docx(out/'page.docx');s.export_pdf(out/'page.pdf')
        shutil.copytree(s.staging_root,out/'native-runtime')
    print(rows)
if __name__=='__main__':main()
