from pathlib import Path
import json, shutil
from skills.WPSComposer import create_document

def main():
    out=Path(__file__).parent
    with create_document('writer',engine='msoffice',visible=False) as s:
        s.add_paragraph('Code ASCII 中文')
        rows=s._execute(['set r to text object of paragraph 1 of boundDoc',
            'set name of font object of r to "Consolas"',
            'set ascii name of font object of r to "Consolas"',
            'set other name of font object of r to "Consolas"',
            'set complex script name of font object of r to "Consolas"',
            'set nativeRows to {{name of font object of r, east asian name of font object of r,ascii name of font object of r,other name of font object of r,complex script name of font object of r}}'])
        (out/'font-slots.json').write_text(json.dumps(rows,ensure_ascii=False))
        s.save_docx(out/'fonts.docx');s.export_pdf(out/'fonts.pdf')
        shutil.copytree(s.staging_root,out/'native-runtime')
    print(rows)
if __name__=='__main__':main()
