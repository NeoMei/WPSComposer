"""Acceptance-only observation: records native host returns and copies job logs.

It neither replaces public APIs nor changes their arguments, results or cleanup.
Only synthetic acceptance workers launched with this directory on PYTHONPATH
load this hook. No installed Python settings are changed.
"""
import os, sys, json, shutil, time
from pathlib import Path

audit = os.environ.get('WPSCOMPOSER_ACCEPTANCE_TRACE')
if audit:
    target = Path(audit)
    target.mkdir(parents=True, exist_ok=True)
    def observe(frame, event, value):
        if event != 'return':
            return
        filename = frame.f_code.co_filename.replace('\\', '/')
        name = frame.f_code.co_name
        row = None
        try:
            if filename.endswith('/msoffice/windows_host.py'):
                if name == 'create_dedicated_composer' and value is not None:
                    row = {'event': 'composer_created', 'word_pid': value.identity.pid,
                           'executable': value.identity.executable, 'binding': value.application_binding,
                           'word_version': str(value.app.Version), 'word_build': str(value.app.Build)}
                elif name == 'close':
                    composer = frame.f_locals['self']
                    row = {'event': 'composer_close_return', 'word_pid': composer.identity.pid,
                           'closed': composer._closed, 'cleanup_error': repr(composer.cleanup_error)}
            elif filename.endswith('/msoffice/windows_runtime.py') and name == '_run_worker':
                operation = frame.f_locals.get('operation')
                if operation and Path(operation).exists():
                    dest = target / 'native-operations' / Path(operation).parent.name / Path(operation).name
                    shutil.copytree(operation, dest)
                    row = {'event': 'worker_return', 'operation': str(operation),
                           'action': frame.f_locals.get('payload', {}).get('action'),
                           'result_returned': value is not None}
            if row:
                row.update(python_pid=os.getpid(), monotonic=time.monotonic())
                with (target / ('trace-' + str(os.getpid()) + '.jsonl')).open('a',encoding='utf-8') as handle:
                    handle.write(json.dumps(row,ensure_ascii=False)+'\n')
        except Exception as exc:
            with (target / 'observer-errors.log').open('a',encoding='utf-8') as handle:
                handle.write(repr(exc)+'\n')
    sys.setprofile(observe)
