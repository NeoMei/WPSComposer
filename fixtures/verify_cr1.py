"""Code-review R1: attach + output= must not rebind the live document."""
import shutil
from pathlib import Path

from skills.WPSComposer import edit
from skills.WPSComposer.scripts._dispatch import _dispatch

FIX = Path(__file__).parent
WORK = FIX / "out" / "cr1_live.docx"
COPY = FIX / "out" / "cr1_copy.docx"
shutil.copy2(FIX / "sample.docx", WORK)
COPY.unlink(missing_ok=True)

app = _dispatch(("KWps.Application", "Wps.Application", "Word.Application"))
app.Visible = -1
doc = app.Documents.Open(str(WORK.resolve()))
orig_name = doc.FullName

r = edit(None, kind="writer", output=str(COPY),
         patches=[{"target": "paragraph:2", "font": {"size": 20}}])
ok = True
ok = ok and r["ok"] and r["saved_path"] == str(COPY.resolve())
ok = ok and COPY.exists()
ok = ok and doc.FullName == orig_name
ok = ok and float(doc.Paragraphs(2).Range.Font.Size) == 20.0
print("PASS attach+output keeps live binding" if ok else "FAIL rebinding/copy issue")
doc.Close(False)
app.Quit()
raise SystemExit(0 if ok else 1)
