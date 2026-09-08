"""Read only a known open synthetic document; does not save/close or Quit."""
import sys,json,hashlib
from pathlib import Path
import pythoncom,win32com.client,win32process
out=Path(__file__).resolve().parent
path=out/'audit-ui-probe.docx'
report=json.loads((out/'com-observations.json').read_text('utf-8'))
pythoncom.CoInitialize()
doc=win32com.client.GetObject(str(path))
hwnd=int(doc.Windows.Item(1).Hwnd)
pid=win32process.GetWindowThreadProcessId(hwnd)[1]
if pid!=report['owned_pid'] or Path(doc.FullName)!=path:
    raise RuntimeError('Readback target is not the verified task document')
text=doc.Content.Text
state={'path':str(path),'pid':pid,'hwnd':hwnd,'saved':bool(doc.Saved),
       'text_sha256':hashlib.sha256(text.encode('utf-8')).hexdigest(),'characters':len(text),
       'audit_marker_present':'WPS-AUDIT-UI-' in text,'end_marker_present':'MSOFFICE-SPIKE-END' in text,
       'tables':[{'rows':doc.Tables.Item(i).Rows.Count,'columns':doc.Tables.Item(i).Columns.Count,
                  'header':bool(doc.Tables.Item(i).Rows.Item(1).HeadingFormat)} for i in range(1,doc.Tables.Count+1)],
       'toc_count':doc.TablesOfContents.Count,'field_count':doc.Fields.Count}
event={'label':sys.argv[1],'transport':'COM read-only verification of existing document','state':state}
report['events'].append(event)
(out/'com-observations.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(event,ensure_ascii=False))
