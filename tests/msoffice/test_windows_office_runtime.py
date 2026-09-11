"""Exercise closed plans, subprocess failures and atomic publication locally."""
from pathlib import Path
from types import SimpleNamespace
import json
import subprocess
import sys
import time
import zipfile

import pytest

from skills.WPSComposer.scripts.generation_plan import GenerationOperation, GenerationPlan, RecordedGeneration


@pytest.fixture
def runtime():
    from skills.WPSComposer.scripts.msoffice import windows_office_runtime
    return windows_office_runtime


def plan(component, operations):
    return GenerationPlan(component, tuple(GenerationOperation(name, args) for name, args in operations))


def office_file(path, fmt):
    family = 'xl' if fmt == 'xlsx' else 'ppt'
    root = 'workbook' if fmt == 'xlsx' else 'presentation'
    namespace = ('http://schemas.openxmlformats.org/spreadsheetml/2006/main' if fmt == 'xlsx'
                 else 'http://schemas.openxmlformats.org/presentationml/2006/main')
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + family + '/' + root + '.xml" ContentType="application/vnd.openxmlformats-officedocument.' + ('spreadsheetml.sheet' if fmt == 'xlsx' else 'presentationml.presentation') + '.main+xml"/></Types>')
        archive.writestr(f'{family}/{root}.xml', f'<{root} xmlns="{namespace}"/>')
        archive.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="' + family + '/' + root + '.xml"/></Relationships>')


def pdf_file(path):
    from pypdf import PdfWriter
    writer = PdfWriter()
    writer.add_blank_page(200, 200)
    with Path(path).open('wb') as stream:
        writer.write(stream)


class JobLock:
    def __init__(self, root):
        self.quarantine_path = root / 'quarantine.json'
    def acquire(self, deadline):
        if self.quarantine_path.exists():
            raise RuntimeError('quarantined')
    def quarantine(self, detail):
        self.quarantine_path.write_text(json.dumps(detail))
    def close(self):
        pass


def local_runtime(runtime, monkeypatch, tmp_path):
    root = tmp_path / 'private'
    root.mkdir()
    monkeypatch.setattr(runtime, '_component_root', lambda component: root)
    monkeypatch.setattr(runtime, 'OfficeJobLock', JobLock)
    monkeypatch.setattr(runtime, 'engine_executable', lambda engine, component: 'C:/Office/native.exe')
    monkeypatch.setattr(runtime, 'sys', SimpleNamespace(platform='win32', executable=sys.executable, argv=sys.argv))
    return root


@pytest.mark.parametrize('component,ops', [
    ('spreadsheet', [('sheet.reset', {}), ('sheet.select', {'index': 9})]),
    ('presentation', [('slide.reset', {}), ('slide.add_image', {'slide': 1, 'imageId': 'missing', 'left': 0, 'top': 0})]),
    ('spreadsheet', [('sheet.reset', {}), ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [['=cmd|"/C calc"!A0']]})]),
])
def test_invalid_plan_rejected_before_native_launch(runtime, monkeypatch, tmp_path, component, ops):
    calls = []
    monkeypatch.setattr(runtime, '_run_worker', lambda *a: calls.append(a))
    with pytest.raises((ValueError, runtime.NativeOfficeError)):
        runtime.generate_recorded(RecordedGeneration(plan(component, ops), ()), tmp_path / ('out.xlsx' if component == 'spreadsheet' else 'out.pptx'))
    assert calls == []


def test_generation_publishes_only_validated_package(runtime, monkeypatch, tmp_path):
    root = local_runtime(runtime, monkeypatch, tmp_path)
    def worker(job, payload, deadline):
        target = job / 'native.xlsx'
        office_file(target, 'xlsx')
        return {'path': str(target), 'cleanup_verified': True}
    monkeypatch.setattr(runtime, '_run_worker', worker)
    output = tmp_path / 'result.xlsx'
    recorded = RecordedGeneration(plan('spreadsheet', [('sheet.reset', {}), ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [['项目', '值'], ['A', 4]]})]), ())
    assert runtime.generate_recorded(recorded, output) == output
    assert zipfile.is_zipfile(output)
    assert not list(root.glob('job-*'))


@pytest.mark.parametrize('fmt,component', [('xlsx', 'spreadsheet'), ('pptx', 'presentation')])
def test_conversion_uses_private_source_copy_and_preserves_original(runtime, monkeypatch, tmp_path, fmt, component):
    root = local_runtime(runtime, monkeypatch, tmp_path)
    source = tmp_path / ('original.' + fmt)
    office_file(source, fmt)
    original = source.read_bytes()
    def worker(job, payload, deadline):
        copy = Path(payload['source'])
        assert copy != source and job in copy.parents and copy.read_bytes() == original
        copy.write_bytes(b'native application modified its task copy')
        target = job / 'export.pdf'
        pdf_file(target)
        return {'path': str(target), 'cleanup_verified': True}
    monkeypatch.setattr(runtime, '_run_worker', worker)
    output = tmp_path / 'out.pdf'
    result = runtime.convert(SimpleNamespace(source=source, output=output, component=component, overwrite=False))
    assert result == output and output.read_bytes().startswith(b'%PDF-')
    assert source.read_bytes() == original


def test_failed_worker_retains_quarantine_without_overwriting_output(runtime, monkeypatch, tmp_path):
    root = local_runtime(runtime, monkeypatch, tmp_path)
    output = tmp_path / 'existing.xlsx'
    output.write_bytes(b'previous deliverable')
    def fail(job, payload, deadline):
        (job / 'partial.xlsx').write_bytes(b'uncertain output')
        raise runtime.NativeOfficeError('NATIVE_OFFICE_TIMEOUT')
    monkeypatch.setattr(runtime, '_run_worker', fail)
    recorded = RecordedGeneration(plan('spreadsheet', [('sheet.reset', {})]), ())
    with pytest.raises(runtime.NativeOfficeError) as exc:
        runtime.generate_recorded(recorded, output, overwrite=True)
    assert exc.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert output.read_bytes() == b'previous deliverable'
    assert (root / 'quarantine.json').exists()
    assert Path(exc.value.staging_path, 'partial.xlsx').read_bytes() == b'uncertain output'


def test_invalid_package_never_replaces_previous_deliverable(runtime, monkeypatch, tmp_path):
    root = local_runtime(runtime, monkeypatch, tmp_path)
    def invalid(job, payload, deadline):
        target = job / 'bad.xlsx'
        target.write_bytes(b'not office')
        return {'path': str(target), 'cleanup_verified': True}
    monkeypatch.setattr(runtime, '_run_worker', invalid)
    output = tmp_path / 'out.xlsx'
    output.write_bytes(b'old')
    with pytest.raises(Exception):
        runtime.generate_recorded(RecordedGeneration(plan('spreadsheet', [('sheet.reset', {})]), ()), output, overwrite=True)
    assert output.read_bytes() == b'old'
    assert not (root / 'quarantine.json').exists()


def test_subprocess_timeout_terminates_only_python_and_retains_diagnostics(runtime, monkeypatch, tmp_path):
    actions = []
    class Child:
        pid = 812
        returncode = None
        def wait(self, timeout):
            actions.append(('wait', timeout))
            if self.returncode is None:
                raise subprocess.TimeoutExpired('native-python-worker', timeout)
        def kill(self):
            actions.append(('kill_python', self.pid))
            self.returncode = -9
    def popen(command, **kwargs):
        assert command[:3] == [sys.executable, '-m', 'skills.WPSComposer.scripts.msoffice.windows_office_runtime']
        return Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', popen)
    with pytest.raises(runtime.NativeOfficeError) as exc:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet', 'action': 'export'}, time.monotonic() + .5)
    assert exc.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert ('kill_python', 812) in actions
    diagnostic = json.loads(Path(exc.value.diagnostic_path).read_text())
    assert diagnostic['office_termination_attempted'] is False


def test_replay_sheet_plan_retains_native_formulas_and_strict_autofit(runtime):
    values = {}
    class Cells:
        def __call__(self, row, col):
            return values.setdefault((row, col), SimpleNamespace(Font=SimpleNamespace(), Interior=SimpleNamespace()))
    class Sheets:
        Count = 1
        def __init__(self):
            self.sheet = SimpleNamespace(Name='Sheet1', Cells=Cells(), UsedRange=SimpleNamespace(Columns=SimpleNamespace(AutoFit=lambda: None)))
        def __call__(self, index):
            return self.sheet
    from skills.WPSComposer.scripts.msoffice.windows_office_host import NativeSheetComposer
    composer = NativeSheetComposer.__new__(NativeSheetComposer)
    composer._doc = SimpleNamespace(Worksheets=Sheets())
    composer._verify_document = lambda: None
    data = plan('spreadsheet', [('sheet.reset', {}), ('sheet.rename', {'index': 1, 'name': '数据'}), ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [['Count', 'Formula'], [2, '=SUM(A2,3)']]}), ('sheet.autofit', {})])
    runtime.execute_plan(composer, data, {})
    assert values[(2, 2)].Formula == '=SUM(A2,3)'
    assert values[(2, 1)].Value == 2
    assert values[(1, 1)].Font.Bold is True
    def fail():
        raise RuntimeError('autofit failed')
    composer._doc.Worksheets.sheet.UsedRange.Columns.AutoFit = fail
    with pytest.raises(RuntimeError, match='autofit failed'):
        runtime.execute_plan(composer, plan('spreadsheet', [('sheet.reset', {}), ('sheet.autofit', {})]), {})


def test_replay_powerpoint_setters_do_not_swallow_native_errors(runtime):
    from skills.WPSComposer.scripts.msoffice.windows_office_host import NativeSlideComposer
    composer = NativeSlideComposer.__new__(NativeSlideComposer)
    composer._doc = SimpleNamespace(Slides=SimpleNamespace(Count=0))
    composer._verify_document = lambda: None
    with pytest.raises(AttributeError):
        runtime.execute_plan(composer, plan('presentation', [('slide.reset', {}), ('slide.set_size', {'width': 960, 'height': 540})]), {})


@pytest.mark.parametrize('formula', ['+cmd|"/C calc"!A0', ' \t=cmd|"/C calc"!A0'])
def test_spreadsheet_formula_prefix_variants_rejected_prelaunch(runtime, formula):
    with pytest.raises(runtime.NativeOfficeError):
        runtime.validate_plan(plan('spreadsheet', [('sheet.reset', {}), ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [[formula]]})]), {})


def test_replay_all_presentation_plan_operations_create_native_objects(runtime, tmp_path):
    class Shape:
        def __init__(self):
            self.TextFrame = SimpleNamespace(TextRange=SimpleNamespace(Text='', Font=SimpleNamespace(Color=SimpleNamespace())))
            self.Fill = SimpleNamespace(ForeColor=SimpleNamespace())
    class Shapes:
        def __init__(self):
            self.Title, self.body = Shape(), Shape()
            self.images = []
            self.tables = []
        def Placeholders(self, index):
            assert index == 2
            return self.body
        def AddPicture(self, path, link, save, left, top, width, height):
            assert link is False and save is True
            shape = SimpleNamespace(path=path, left=left, top=top, Width=width, Height=height)
            self.images.append(shape)
            return shape
        def AddTable(self, rows, cols, left, top, width, height):
            cells = {(r, c): SimpleNamespace(Shape=Shape()) for r in range(1, rows + 1) for c in range(1, cols + 1)}
            table = SimpleNamespace(Cell=lambda r, c: cells[(r, c)], cells=cells)
            self.tables.append(table)
            return SimpleNamespace(Table=table)
    class Slides:
        def __init__(self):
            self.items = []
        @property
        def Count(self):
            return len(self.items)
        def Add(self, index, layout):
            slide = SimpleNamespace(Shapes=Shapes(), layout=layout)
            self.items.insert(index - 1, slide)
            return slide
        def __call__(self, index):
            return self.items[index - 1]
    from skills.WPSComposer.scripts.msoffice.windows_office_host import NativeSlideComposer
    from skills.WPSComposer.scripts.recording_composers import RecordingSlideComposer
    from skills.WPSComposer.scripts.design_presets import PRESETS
    image = tmp_path / 'image.png'
    image.write_bytes(b'private resource boundary')
    with RecordingSlideComposer() as recording:
        recording.set_slide_size(960, 540)
        recording.apply_design_preset(PRESETS['business'])
        recording.add_title_slide('业务标题', '副标题')
        recording.add_section_slide('章节')
        recording.add_bullets_slide('数据', ['第一项', '第二项'])
        recording.add_blank_slide()
        recording.add_image(4, image, 10, 20, 100, 80)
        recording.add_table(4, 2, 2, 20, 120, 200, 100, [['项目', '值'], ['A', 3]])
        recorded = recording.save_pptx('unused.pptx')
    composer = NativeSlideComposer.__new__(NativeSlideComposer)
    composer._verify_document = lambda: None
    composer._doc = SimpleNamespace(Slides=Slides(), PageSetup=SimpleNamespace(),
        SlideMaster=SimpleNamespace(Background=SimpleNamespace(Fill=SimpleNamespace(ForeColor=SimpleNamespace(), Solid=lambda: None))))
    runtime.execute_plan(composer, recorded.plan, {'image-1': image})
    slides = composer._doc.Slides.items
    assert len(slides) == 4 and [s.layout for s in slides] == [1, 11, 2, 12]
    assert slides[0].Shapes.Title.TextFrame.TextRange.Text == '业务标题'
    assert slides[0].Shapes.body.TextFrame.TextRange.Text == '副标题'
    assert slides[2].Shapes.body.TextFrame.TextRange.Text == '第一项\r第二项'
    assert Path(slides[3].Shapes.images[0].path.replace('\\', '/')) == image
    assert slides[3].Shapes.tables[0].cells[(2, 2)].Shape.TextFrame.TextRange.Text == '3'
    assert composer._doc.PageSetup.SlideWidth == 960


def test_replay_sheet_add_select_and_column_width_remain_on_owned_sheets(runtime):
    def sheet(name):
        widths = {}
        cells = {}
        return SimpleNamespace(Name=name, widths=widths, cells=cells,
            Columns=lambda name: widths.setdefault(name, SimpleNamespace()),
            Cells=lambda row, col: cells.setdefault((row, col), SimpleNamespace(Font=SimpleNamespace(), Interior=SimpleNamespace())))
    class Sheets:
        def __init__(self):
            self.items = [sheet('Sheet1')]
        @property
        def Count(self):
            return len(self.items)
        def __call__(self, index):
            return self.items[index - 1]
        def Add(self, before, after):
            assert before is None
            current = sheet('new')
            self.items.insert(self.items.index(after) + 1, current)
            return current
    from skills.WPSComposer.scripts.msoffice.windows_office_host import NativeSheetComposer
    composer = NativeSheetComposer.__new__(NativeSheetComposer)
    composer._doc = SimpleNamespace(Worksheets=Sheets())
    composer._verify_document = lambda: None
    runtime.execute_plan(composer, plan('spreadsheet', [
        ('sheet.reset', {}), ('sheet.rename', {'index': 1, 'name': 'First'}),
        ('sheet.add', {'name': 'Second'}), ('sheet.set_column_width', {'column': 'B', 'width': 24}),
        ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [[7]]}),
        ('sheet.select', {'index': 1}), ('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [[8]]})]), {})
    first, second = composer._doc.Worksheets.items
    assert [first.Name, second.Name] == ['First', 'Second']
    assert first.cells[(1, 1)].Value == 8 and second.cells[(1, 1)].Value == 7
    assert second.widths['B'].ColumnWidth == 24 and not first.widths


def test_windows_job_lock_blocks_second_job_and_retains_quarantine(runtime, monkeypatch, tmp_path):
    import errno
    held = set()
    def locking(fd, mode, length):
        assert length == 1
        if mode == 1:
            if held:
                raise OSError(errno.EACCES, 'already locked')
            held.add(fd)
        else:
            held.remove(fd)
    monkeypatch.setitem(sys.modules, 'msvcrt', SimpleNamespace(locking=locking, LK_NBLCK=1, LK_UNLCK=2))
    first, second = runtime.OfficeJobLock(tmp_path), runtime.OfficeJobLock(tmp_path)
    first.acquire(time.monotonic() + 1)
    with pytest.raises(runtime.NativeOfficeError) as exc:
        second.acquire(time.monotonic() + .01)
    assert exc.value.code == 'NATIVE_OFFICE_TIMEOUT'
    first.quarantine({'status': 'uncertain'})
    first.close()
    with pytest.raises(runtime.NativeOfficeError) as exc:
        second.acquire(time.monotonic() + 1)
    assert exc.value.code == 'NATIVE_OFFICE_QUARANTINED'
    assert not held
