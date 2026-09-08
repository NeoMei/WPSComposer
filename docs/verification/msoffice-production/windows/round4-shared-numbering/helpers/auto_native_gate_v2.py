import sys, json, subprocess, hashlib, zipfile, traceback, winreg
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer.scripts.office_engines import engine_executable, resolve_engine
from skills.WPSComposer import generate, convert_to_pdf
from fixtures.verify_msoffice_production import fixture_content, inspect_docx, inspect_pdf
root = Path('build/msoffice-production/auto-02').resolve(); root.mkdir(exist_ok=False)
report = {'checks':{}, 'registration':[], 'status':'RUNNING'}
def persist(): (root / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
def processes():
    import win32com.client
    service = win32com.client.GetObject('winmgmts:')
    return sorted(int(p.ProcessId) for p in service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='WINWORD.EXE' OR Name='wps.exe'"))
try:
    before = processes()
    report['detected'] = {engine:engine_executable(engine) for engine in ['wps', 'msoffice']}
    report['resolved'] = resolve_engine('auto', 'writer')
    report['checks']['read_only_detection_starts_no_office_process'] = processes() == before
    report['checks']['auto_selects_real_wps'] = report['resolved'] == 'wps' and Path(report['detected']['wps']).name.lower() == 'wps.exe'
    for label, view in [('default',0), ('32',winreg.KEY_WOW64_32KEY), ('64',winreg.KEY_WOW64_64KEY)]:
        row = {'view':label}
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'KWps.Application\CLSID', 0, winreg.KEY_READ | view) as key:
                row['clsid'] = winreg.QueryValueEx(key, None)[0]
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + row['clsid'] + r'\LocalServer32', 0, winreg.KEY_READ | view) as key:
                row['local_server'] = winreg.QueryValueEx(key, None)[0]
        except OSError as exc: row['error'] = str(exc)
        report['registration'].append(row)
    persist()
    source, spec = fixture_content('smoke'); (root/'source.md').write_text(source, encoding='utf-8')
    output = root/'generated.docx'
    generate(source, source_is_text=True, format='docx', output=str(output), engine='auto', timeout=300, open_result=False)
    before_hash = hashlib.sha256(output.read_bytes()).hexdigest()
    convert_to_pdf(str(output), str(root/'converted.pdf'), engine='auto', timeout=300, open_result=False)
    report['docx'] = inspect_docx(output, spec); report['pdf'] = inspect_pdf(root/'converted.pdf', spec)
    with zipfile.ZipFile(output) as archive:
        report['application_xml'] = archive.read('docProps/app.xml').decode()
    report['checks']['wps_artifact_provenance'] = 'WPS Office' in report['application_xml']
    report['checks']['conversion_preserves_docx'] = hashlib.sha256(output.read_bytes()).hexdigest() == before_hash
    checks = list(report['checks'].values()) + list(report['docx']['checks'].values()) + list(report['pdf']['checks'].values())
    report['status'] = 'PASS' if all(checks) else 'FAIL'
except Exception as exc:
    report['status'] = 'FAIL'; report['error'] = repr(exc); report['traceback'] = traceback.format_exc()
persist(); print(report['status'])
raise SystemExit(0 if report['status']=='PASS' else 1)
