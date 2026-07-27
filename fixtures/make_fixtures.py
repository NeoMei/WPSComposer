import base64
from pathlib import Path

from skills.WPSComposer import generate

FIX = Path(__file__).parent
PNG_1X1 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk"
    "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)


def main():
    (FIX / "logo.png").write_bytes(base64.b64decode(PNG_1X1))

    docx_md = FIX / "_docx.md"
    docx_md.write_text(
        "# Sample Title\n\n"
        "## Section One\n\n"
        "First paragraph of the sample document.\n\n"
        "Second paragraph with some more text content.\n\n"
        "Third paragraph completes the trio.\n\n"
        "| Name | Value |\n|------|-------|\n| alpha | 1 |\n| beta | 2 |\n\n"
        "## Section Two\n\n"
        "Closing paragraph for the sample.\n",
        encoding="utf-8",
    )
    generate(str(docx_md), format="docx", output=str(FIX / "sample.docx"))

    pptx_md = FIX / "_pptx.md"
    pptx_md.write_text(
        "# Deck Title\n\nSubtitle text\n\n"
        "---\n\n# Slide Two\n\n- bullet one\n- bullet two\n\n"
        "---\n\n# Slide Three\n\nFinal slide body text.\n",
        encoding="utf-8",
    )
    generate(str(pptx_md), format="pptx", output=str(FIX / "sample.pptx"))

    xlsx_md = FIX / "_xlsx.md"
    xlsx_md.write_text(
        "| Item | Q1 | Q2 |\n|------|----|----|\n"
        "| apples | 10 | 20 |\n| oranges | 15 | 25 |\n| pears | 12 | 18 |\n",
        encoding="utf-8",
    )
    generate(str(xlsx_md), format="xlsx", output=str(FIX / "sample.xlsx"))

    # blank docx for replay tests
    blank_md = FIX / "_blank.md"
    blank_md.write_text("# Blank\n\nplaceholder\n", encoding="utf-8")
    generate(str(blank_md), format="docx", output=str(FIX / "blank.docx"))

    for stub in FIX.glob("_*.md"):
        stub.unlink()
    print("fixtures created:", sorted(p.name for p in FIX.iterdir()))


if __name__ == "__main__":
    main()
