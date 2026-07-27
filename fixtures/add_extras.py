from pathlib import Path

from skills.WPSComposer.scripts.writer import WriterComposer
from skills.WPSComposer.scripts.sheet import SheetComposer

FIX = Path(__file__).parent


def main():
    with WriterComposer.open_document(str(FIX / "sample.docx")) as writer:
        writer._doc.Shapes.AddTextbox(1, 100, 100, 200, 50).TextFrame.TextRange.Text = "fixture textbox"
        writer.save(str(FIX / "sample.docx"))

    with SheetComposer.open_document(str(FIX / "sample.xlsx")) as sheet:
        sheet.ws = sheet._doc.Worksheets(1)
        sheet.add_chart(chart_type=51, left=300, top=10, width=400, height=300,
                        source_range="A1:C4", title="Sales")
        sheet.save(str(FIX / "sample.xlsx"))

    print("extras added")


if __name__ == "__main__":
    main()
