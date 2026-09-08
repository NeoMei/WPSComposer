"""Two real chapters per numbering scheme, including native insert/delete/reopen."""
import sys, json, traceback, re, hashlib
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import pythoncom, win32com.client, fitz
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer
from skills.WPSComposer.scripts.writer import WriterComposer

engine = sys.argv[1]
root = Path('build/msoffice-production/schemes-' + engine + '-01').resolve()
root.mkdir(exist_ok=False)
report = {'engine': engine, 'cases': {}, 'source_sha256': {}}
for filename in ['writer.py', 'msoffice/windows_host.py']:
    report['source_sha256'][filename] = hashlib.sha256((Path('skills/WPSComposer/scripts') / filename).read_bytes()).hexdigest()
def persist():
    (root / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
def expected(scheme, chapter):
    chinese = {1:'一', 2:'二', 3:'三'}[chapter]
    if scheme == 'decimal':
        return [str(chapter), f'{chapter}.1', f'{chapter}.1.1', f'{chapter}.1.1.1']
    if scheme == 'chinese-formal':
        return [f'第{chinese}章', '第一节', '一、', '（一）']
    return [f'第{chinese}章', f'{chapter}.1', f'{chapter}.1.1', '关键工法01：']
for scheme in ['decimal', 'chinese-formal', 'hybrid-bid']:
    case = report['cases'][scheme] = {'status':'RUNNING', 'checks':{}, 'states':{}}
    folder = root / scheme
    folder.mkdir()
    c = doc = None
    try:
        if engine == 'word':
            c = create_dedicated_composer(str(folder)); doc = c.doc
            case['pid'] = c.identity.pid; case['binding'] = c.application_binding
        else:
            pythoncom.CoInitialize(); app = win32com.client.DispatchEx('KWps.Application')
            assert Path(app.Path).resolve() == Path('C:/Users/1/AppData/Local/12.1.0.28505/office6').resolve()
            doc = app.Documents.Add()
            assert doc.Application._oleobj_.QueryInterface(pythoncom.IID_IUnknown) == app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
            c = WriterComposer.__new__(WriterComposer); c._app = app; c._doc = doc
            case['application'] = {'name':app.Name, 'path':app.Path, 'version':app.Version}
        def snapshot(label):
            rows = []
            for i in range(1, doc.Paragraphs.Count + 1):
                r = doc.Paragraphs(i).Range
                text = r.Text.strip()
                if text:
                    rows.append({'text':text, 'number':str(r.ListFormat.ListString)})
            case['states'][label] = rows; persist()
            return [row['number'] for row in rows]
        def save(name):
            if engine == 'word': c.save_docx(folder / name)
            else: doc.SaveAs(str(folder / name), 12)
        for chapter in range(1, 3):
            for level in range(1, 5):
                c.add_heading_level_native(f'Chapter {chapter} Level {level}', level, numbering=True, scheme=scheme)
        case['checks']['two_chapters'] = snapshot('before') == expected(scheme, 1) + expected(scheme, 2)
        template = c._wpsc_heading_templates[scheme]
        case['levels'] = [{'level':i, 'format':template.ListLevels(i).NumberFormat, 'style':template.ListLevels(i).NumberStyle, 'reset':template.ListLevels(i).ResetOnHigher} for i in range(1, 5)]
        save('before.docx')
        marker = 'Inserted Chapter\r'
        doc.Range(0, 0).InsertBefore(marker)
        doc.Range(0, len(marker)).Style = doc.Styles(-2)
        c.update_fields()
        case['checks']['insert_renumbers_all_levels'] = snapshot('inserted') == [expected(scheme, 1)[0]] + expected(scheme, 2) + expected(scheme, 3)
        save('inserted.docx')
        doc.Paragraphs(1).Range.Delete(); c.update_fields()
        case['checks']['delete_restores_all_levels'] = snapshot('deleted') == expected(scheme, 1) + expected(scheme, 2)
        save('restored.docx')
        if engine == 'word':
            c.open_owned_document(folder / 'restored.docx', read_only=False); doc = c.doc
            c.export_pdf(folder / 'restored.pdf')
        else:
            doc.Close(SaveChanges=0)
            doc = app.Documents.Open(str(folder / 'restored.docx')); c._doc = doc
            doc.ExportAsFixedFormat(str(folder / 'restored.pdf'), 17)
        case['checks']['reopen_retains_all_levels'] = snapshot('reopened') == expected(scheme, 1) + expected(scheme, 2)
        with fitz.open(folder / 'restored.pdf') as pdf:
            case['pdf_text'] = '\n'.join(page.get_text() for page in pdf)
            compact = re.sub(r'\s+', '', case['pdf_text'])
            case['checks']['pdf_displays_expected_numbers'] = all(
                re.sub(r'\s+', '', number + f'Chapter {chapter} Level {level}') in compact
                for chapter in range(1, 3) for level, number in enumerate(expected(scheme, chapter), 1))
            page = pdf[0]; page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2)).save(str(folder / 'page-1.png'))
        case['status'] = 'PASS' if all(case['checks'].values()) else 'FAIL'
    except Exception as exc:
        case['status'] = 'FAIL'; case['error'] = repr(exc); case['traceback'] = traceback.format_exc()
    finally:
        if c is not None:
            if engine == 'word': c.close(save_changes=False); case['closed'] = c._closed
            elif doc is not None: doc.Close(SaveChanges=0); case['owned_document_closed'] = True; case['quit_attempted'] = False
        persist()
    print(scheme, case['status'], flush=True)
report['status'] = 'PASS' if all(c['status'] == 'PASS' for c in report['cases'].values()) else 'FAIL'
persist()
raise SystemExit(0 if report['status'] == 'PASS' else 1)
