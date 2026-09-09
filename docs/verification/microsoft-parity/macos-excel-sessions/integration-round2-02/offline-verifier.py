from pathlib import Path
import hashlib
import json

from openpyxl import load_workbook
from pypdf import PdfReader

root = Path("docs/verification/microsoft-parity/macos-excel-sessions/integration-round2-02")
report = json.loads((root / "report.json").read_text())
assert report["passed"] and len(report["checks"]) == 8 and all(report["checks"].values())
for relative, expected in report["source_hashes"].items():
    current = Path(relative)
    retained = root / "source" / relative
    assert hashlib.sha256(current.read_bytes()).hexdigest() == expected
    assert hashlib.sha256(retained.read_bytes()).hexdigest() == expected
for name, expected in report["artifact_hashes"].items():
    assert hashlib.sha256((root / name).read_bytes()).hexdigest() == expected
assert report["artifact_hashes"]["bound.xlsx"] == report["artifact_hashes"]["concurrent.xlsx"]
assert report["artifact_hashes"]["bound.xlsx"] != report["artifact_hashes"]["recovery.xlsx"]

values = {}
for name in ("bound.xlsx", "concurrent.xlsx", "recovery.xlsx"):
    workbook = load_workbook(root / name, read_only=True, data_only=False)
    try:
        sheet = workbook.worksheets[0]
        values[name] = {"A30": sheet["A30"].value, "A31": sheet["A31"].value}
    finally:
        workbook.close()
assert values["bound.xlsx"] == values["concurrent.xlsx"] == {
    "A30": "PUBLISHED BEFORE CONFLICT", "A31": None,
}
assert values["recovery.xlsx"] == {
    "A30": "PUBLISHED BEFORE CONFLICT", "A31": "UNPUBLISHED RECOVERY EDIT",
}

pdf = PdfReader(str(root / "recovery.pdf"), strict=True)
assert len(pdf.pages) >= 1
pdf_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
assert "PUBLISHED BEFORE CONFLICT" in pdf_text
assert "UNPUBLISHED RECOVERY EDIT" in pdf_text

before = json.loads((root / "inventory-before.json").read_text())
final = json.loads((root / "inventory-final.json").read_text())
during = json.loads((root / "inventory-after-failed-close.json").read_text())
assert before == final
assert any(
    item["name"] == report["owned_name"] and item["path"] == report["owned_path"]
    for item in during
)
close_scripts = [
    path.name for path in (root / "native-job").glob("step-*.applescript")
    if "close ownedBook saving no" in path.read_text()
]
assert close_scripts == [report["close_script"]]
print(json.dumps({"checks":len(report["checks"]),"sources":len(report["source_hashes"]),"artifacts":len(report["artifact_hashes"]),"cells":values,"pdf_pages":len(pdf.pages),"inventory_restored":before==final,"close_scripts":close_scripts}, ensure_ascii=False))
