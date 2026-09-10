"""Fresh candidate public API quarantine check; no removal or recovery."""
import datetime, hashlib, importlib.util, json, sys, traceback
from pathlib import Path
root=Path('D:/wpsc70')
out=Path(__file__).parent
sys.path.insert(0,str(root))
import skills.WPSComposer as api
from skills.WPSComposer.scripts.msoffice.windows_office_runtime import _component_root
spec=importlib.util.spec_from_file_location('candidate_gate',root/'fixtures/microsoft_parity/evidence_gate.py')
gate=importlib.util.module_from_spec(spec);spec.loader.exec_module(gate)

def main():
    report=dict(candidate='70d3d86ff6b03945dacccffd1688922830ad6879',
        time=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        execution='public API startup preflight; not native acceptance',
        imported_api=api.__file__,python=sys.executable,source_digest=gate.source_digest(root),cases={})
    for kind,component in [('writer','writer'),('sheet','spreadsheet'),('slide','presentation')]:
        marker=_component_root(component)/'native-office.quarantine.json'
        raw=marker.read_bytes()
        (out/(component+'-quarantine.json')).write_bytes(raw)
        before_dirs=sorted(p.name for p in marker.parent.glob('session-*'))
        case=dict(component=component,quarantine=str(marker),native_executed=False,ui_executed=False)
        try:
            session=api.create_document(kind,engine='msoffice')
            case.update(status='unexpectedly_started',passed=False)
            # Do not make subsequent calls if isolation unexpectedly changes.
            raise RuntimeError('Unexpectedly started despite quarantine; preserve returned session')
        except BaseException as exc:
            case.update(status='blocked',error_type=type(exc).__name__,error=str(exc),code=getattr(exc,'code',None))
            (out/(kind+'-preflight-error.txt')).write_text(traceback.format_exc(),encoding='utf-8')
        case['quarantine_unchanged']=marker.read_bytes()==raw
        case['no_session_directory_created']=before_dirs==sorted(p.name for p in marker.parent.glob('session-*'))
        case['quarantine_sha256']=hashlib.sha256(raw).hexdigest()
        report['cases'][kind]=case
    report['source_digest_after']=gate.source_digest(root)
    report['native_acceptance_passed']=False
    (out/'preflight-report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))
    return 2

if __name__=='__main__':
    raise SystemExit(main())

