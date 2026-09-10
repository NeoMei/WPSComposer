from __future__ import annotations

from dataclasses import replace
import json
from pathlib import Path
import re
import subprocess

import pytest

from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity, OSATransportError


def api():
    from skills.WPSComposer.scripts.msoffice import macos_excel_process
    return macos_excel_process


def identity():
    return ExcelProcessIdentity(4242, 100, 250,
        '/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel', 'com.microsoft.Excel')


class Clock:
    now = 1.0
    def __call__(self):
        return self.now


class Launcher:
    def __init__(self, clock):
        self.clock = clock
        self.current = identity()
        self.before = {9000}
        self.ready = True
        self.exited = False
        self.waits = []
        self.launches = 0
        self.record = None

    def existing_pids(self):
        return self.before

    def launch(self, app_path, deadline):
        self.launches += 1
        self.record = api().LaunchedExcel(identity(), 100.00025, object())
        return self.record

    def snapshot(self, pid):
        assert pid == 4242
        return self.current

    def finished_launching(self, record):
        assert record is self.record
        return self.ready

    def has_terminated(self, record):
        assert record is self.record
        return self.exited

    def wait_until(self, deadline):
        self.waits.append(deadline)
        self.clock.now = deadline


def book(path='/tmp/owned.xlsx', *, pristine=False):
    return dict(name=Path(path).name, path=str(Path(path).parent), full_name=path,
                saved=True, pristine=pristine)


class Transport:
    def __init__(self, launcher, books=()):
        self.launcher = launcher
        self.books = list(books)
        self.events = []
        self.failure = None
        self.bad_nonce = False
        self.exit_on_quit = True
        self.after_close = []

    def run(self, script_path, deadline):
        source = script_path.read_text()
        operation = re.search(r'-- owner operation: (\w+)', source).group(1)
        self.events.append(operation)
        if self.failure:
            raise self.failure
        nonce = re.search(r'set ownerNonce to "([a-f0-9]+)"', source).group(1)
        reply = {'nonce': 'wrong' if self.bad_nonce else nonce}
        if operation == 'inventory':
            reply.update(version='16.101', workbooks=self.books)
        elif operation in ('bootstrap', 'close'):
            self.books = self.after_close
            reply['empty'] = not self.books
        elif operation == 'quit':
            reply['quit_sent'] = True
            if self.exit_on_quit:
                self.launcher.current = None
                self.launcher.exited = True
        else:
            raise AssertionError(operation)
        return subprocess.CompletedProcess([], 0, json.dumps(reply), '')


def setup_owner(*, books=(), **options):
    clock = Clock()
    launcher = Launcher(clock)
    transport = Transport(launcher, books)
    owner = api().PrivateExcelProcessOwner(launcher=launcher,
        transport_factory=lambda exact: transport, clock=clock, **options)
    return owner, launcher, transport, clock


def test_start_binds_returned_instance_and_requires_empty_inventory():
    owner, launcher, transport, _ = setup_owner()
    owner.start(deadline=5)
    assert owner.identity == identity()
    assert owner.launch_date == 100.00025
    assert owner.is_private_owned
    assert owner.state == 'ready'
    assert transport.events == ['inventory']
    assert launcher.launches == 1


def test_reused_existing_process_is_never_contacted_or_quit():
    owner, launcher, transport, _ = setup_owner()
    launcher.before.add(4242)
    with pytest.raises(api().ExcelProcessError, match='EXCEL_LAUNCH_REUSED_PROCESS'):
        owner.start(deadline=5)
    assert transport.events == []
    assert owner.state == 'quarantined'
    assert not owner.is_private_owned


def test_readiness_waits_by_deadline_and_records_recovery_without_guessing_process():
    owner, launcher, transport, clock = setup_owner()
    launcher.ready = False
    with pytest.raises(api().ExcelProcessError, match='EXCEL_DEADLINE'):
        owner.start(deadline=1.3)
    assert clock.now == 1.3
    assert launcher.waits and max(launcher.waits) <= 1.3
    assert transport.events == []
    assert owner.recovery['identity']['start_microseconds'] == 250
    assert owner.state == 'quarantined'


@pytest.mark.parametrize('change', [dict(start_seconds=101), dict(executable='/tmp/Microsoft Excel'), dict(bundle_id='foreign')])
def test_identity_drift_before_readiness_never_sends(change):
    owner, launcher, transport, _ = setup_owner()
    # Launch and lookup must agree with the returned handle as well as birth token.
    original_launch = launcher.launch
    def launch(*args):
        result = original_launch(*args)
        launcher.current = replace(identity(), **change)
        return result
    launcher.launch = launch
    with pytest.raises((api().ExcelProcessError, OSATransportError)):
        owner.start(deadline=5)
    assert transport.events == []
    assert owner.state == 'quarantined'


@pytest.mark.parametrize('books', [[book()], [book('/tmp/a.xlsx'), book('/tmp/b.xlsx')],
    [dict(name='Book1', path='', full_name='Book1', saved=False, pristine=True)]])
def test_foreign_or_unproven_bootstrap_inventory_is_not_closed(books):
    owner, _, transport, _ = setup_owner(books=books)
    with pytest.raises(api().ExcelProcessError, match='EXCEL_FOREIGN_WORKBOOKS'):
        owner.start(deadline=5)
    assert transport.events == ['inventory']
    assert owner.recovery['inventory'] == books


def test_only_pristine_initial_book_is_normalized_and_empty_inventory_rechecked():
    bootstrap = dict(name='Book1', path='', full_name='Book1', saved=True, pristine=True)
    owner, _, transport, _ = setup_owner(books=[bootstrap])
    owner.start(deadline=5)
    assert transport.events == ['inventory', 'bootstrap', 'inventory']
    assert owner.is_private_owned


def test_bad_readiness_nonce_quarantines_without_normalization():
    owner, _, transport, _ = setup_owner()
    transport.bad_nonce = True
    with pytest.raises(api().ExcelProcessError, match='EXCEL_BAD_ACK'):
        owner.start(deadline=5)
    assert owner.state == 'quarantined'
    assert transport.events == ['inventory']


def test_close_exact_owned_path_then_quit_and_acknowledge_birth_disappearance():
    owner, launcher, transport, _ = setup_owner()
    owner.start(deadline=5)
    transport.books = [book()]
    owner.reserve_workbook(Path('/tmp/owned.xlsx'))
    owner.claim_workbook(Path('/tmp/owned.xlsx'), deadline=5)
    owner.close(deadline=5)
    assert transport.events == ['inventory', 'inventory', 'close', 'inventory', 'quit']
    assert launcher.exited
    assert owner.state == 'closed'
    assert not owner.is_private_owned
    assert owner.recovery is None


def test_foreign_book_after_close_blocks_quit_and_keeps_recovery_path():
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    transport.books = [book()]
    owner.reserve_workbook(Path('/tmp/owned.xlsx'))
    owner.claim_workbook(Path('/tmp/owned.xlsx'), deadline=5)
    transport.after_close = [book('/tmp/user.xlsx')]
    with pytest.raises(api().ExcelProcessError, match='EXCEL_BAD_ACK'):
        owner.close(deadline=5)
    assert 'quit' not in transport.events
    assert owner.recovery['owned_path'] == '/tmp/owned.xlsx'


def test_quit_timeout_quarantines_and_does_not_kill_or_retry():
    owner, launcher, transport, clock = setup_owner()
    owner.start(deadline=5)
    transport.exit_on_quit = False
    with pytest.raises(api().ExcelProcessError, match='EXCEL_DEADLINE'):
        owner.close(deadline=1.2)
    assert clock.now == 1.2
    assert launcher.current == identity()
    assert transport.events.count('quit') == 1
    assert owner.recovery['identity']['pid'] == 4242
    with pytest.raises(api().ExcelProcessError, match='EXCEL_OWNER_UNAVAILABLE'):
        owner.close(deadline=5)


def test_pid_lookup_failure_is_not_exit_acknowledgment():
    owner, launcher, transport, _ = setup_owner()
    owner.start(deadline=5)
    transport.exit_on_quit = False
    original_run = transport.run
    def run(*args):
        result = original_run(*args)
        if transport.events[-1] == 'quit':
            launcher.current = None
        return result
    transport.run = run
    with pytest.raises(api().ExcelProcessError, match='EXCEL_DEADLINE'):
        owner.close(deadline=1.1)
    assert owner.state == 'quarantined'


def test_uncertain_transport_error_retains_owned_path_and_prevents_further_commands(tmp_path):
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    transport.books = [book()]
    owner.reserve_workbook(Path('/tmp/owned.xlsx'))
    owner.claim_workbook(Path('/tmp/owned.xlsx'), deadline=5)
    script = tmp_path / 'mutation.applescript'
    script.write_text('-- owner operation: mutate')
    transport.failure = OSATransportError('OSA_TIMEOUT', outcome_uncertain=True)
    with pytest.raises(OSATransportError):
        owner.run(script, deadline=5)
    assert owner.recovery['owned_path'] == '/tmp/owned.xlsx'
    assert owner.recovery['outcome_uncertain'] is True
    with pytest.raises(api().ExcelProcessError, match='EXCEL_OWNER_UNAVAILABLE'):
        owner.run(script, deadline=5)
    assert transport.events.count('mutate') == 1


def test_reserved_path_survives_failure_before_workbook_claim(tmp_path):
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    owner.reserve_workbook(Path('/tmp/planned.xlsx'))
    script = tmp_path / 'open.applescript'
    script.write_text('-- owner operation: open')
    transport.failure = OSATransportError('OSA_TIMEOUT', outcome_uncertain=True)
    with pytest.raises(OSATransportError):
        owner.run(script, deadline=5)
    assert owner.recovery['owned_path'] == '/tmp/planned.xlsx'


@pytest.mark.parametrize('date', [None, float('nan')])
def test_ambiguous_launch_date_is_rejected_before_any_command(date):
    owner, launcher, transport, _ = setup_owner()
    original_launch = launcher.launch
    def launch(*args):
        launcher.record = replace(original_launch(*args), launch_date=date)
        return launcher.record
    launcher.launch = launch
    with pytest.raises(api().ExcelProcessError, match='EXCEL_LAUNCH_DATE_MISMATCH'):
        owner.start(deadline=5)
    assert transport.events == []


def test_partial_native_launch_failure_preserves_returned_pid_without_adopting_it():
    owner, launcher, transport, _ = setup_owner()
    def launch(*args):
        launcher.recovery_candidate = {'pid': 4242, 'bundle_id': 'com.microsoft.Excel'}
        raise api().ExcelProcessError('EXCEL_LAUNCH_IDENTITY_UNAVAILABLE')
    launcher.launch = launch
    with pytest.raises(api().ExcelProcessError):
        owner.start(deadline=5)
    assert owner.recovery['launch_candidate']['pid'] == 4242
    assert owner.identity is None
    assert transport.events == []


def test_identity_replaced_after_successful_event_quarantines_session(tmp_path):
    owner, launcher, transport, _ = setup_owner()
    owner.start(deadline=5)
    script = tmp_path / 'mutate.applescript'
    owner.reserve_workbook(Path('/tmp/planned.xlsx'))
    script.write_text('mutation source')
    def run(*args):
        launcher.current = replace(identity(), start_seconds=200)
        return subprocess.CompletedProcess([], 0, '{}', '')
    transport.run = run
    with pytest.raises(OSATransportError, match='OSA_PROCESS_IDENTITY_CHANGED'):
        owner.run(script, deadline=5)
    assert owner.state == 'quarantined'


def test_quit_pid_replacement_acknowledges_old_exit_without_targeting_replacement():
    owner, launcher, transport, _ = setup_owner()
    owner.start(deadline=5)
    original_run = transport.run
    def run(*args):
        result = original_run(*args)
        if transport.events[-1] == 'quit':
            launcher.current = replace(identity(), start_seconds=200)
        return result
    transport.run = run
    owner.close(deadline=5)
    assert owner.state == 'closed'
    assert launcher.current.start_seconds == 200
    assert transport.events == ['inventory', 'inventory', 'quit']


class CocoaRuntime:
    """Objective-C boundary double: no frameworks or application processes."""
    def __init__(self):
        self.options = None
        self.calls = []
        self.launch_result = 'returned-app'
        self.date = 100.00025

    def objc_class(self, name):
        return name

    def ns_string(self, value):
        return value

    def object_text(self, value):
        return value

    def message(self, receiver, selector, args=(), types=(), result=None):
        self.calls.append((receiver, selector, args, types, result))
        values = {'new': 'pool', 'sharedWorkspace': 'workspace', 'dictionary': 'config',
            'processIdentifier': 4242, 'bundleIdentifier': 'com.microsoft.Excel',
            'executableURL': 'executable-url', 'path': identity().executable,
            'launchDate': 'date', 'timeIntervalSince1970': self.date,
            'retain': receiver, 'release': None, 'drain': None,
            'isFinishedLaunching': True, 'isTerminated': False}
        if selector == 'fileURLWithPath:':
            assert args == ('/Applications/Microsoft Excel.app',)
            return 'app-url'
        if selector == 'launchApplicationAtURL:options:configuration:error:':
            assert receiver == 'workspace'
            assert args[0] == 'app-url'
            assert args[2] == 'config'
            self.options = args[1]
            return self.launch_result
        if selector in values:
            return values[selector]
        raise AssertionError((receiver, selector, args))


def native_launcher(runtime):
    launcher = api().AppKitExcelLauncher.__new__(api().AppKitExcelLauncher)
    launcher.runtime = runtime
    launcher._lookup = type('Lookup', (), {'snapshot': lambda self, pid, **kwargs: identity()})()
    launcher._clock = lambda: 1.0
    launcher._retained = []
    launcher._applications = {}
    launcher.recovery_candidate = None
    return launcher


def test_appkit_launch_uses_returned_handle_new_instance_async_and_allows_activation():
    runtime = CocoaRuntime()
    launcher = native_launcher(runtime)
    result = launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    assert result.handle == 'returned-app'
    assert result.identity == identity()
    assert result.launch_date == 100.00025
    assert runtime.options & 0x80000  # NSWorkspaceLaunchNewInstance
    assert runtime.options & 0x10000  # NSWorkspaceLaunchAsync
    assert not runtime.options & 0x200  # NSWorkspaceLaunchWithoutActivation
    assert ('returned-app', 'retain') in [(r, s) for r, s, *_ in runtime.calls]
    assert ('pool', 'drain') in [(r, s) for r, s, *_ in runtime.calls]


def test_appkit_nil_result_fails_without_pid_discovery_and_drains_pool():
    runtime = CocoaRuntime()
    runtime.launch_result = None
    launcher = native_launcher(runtime)
    with pytest.raises(api().ExcelProcessError, match='EXCEL_LAUNCH_FAILED'):
        launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    assert 'processIdentifier' not in [s for _, s, *_ in runtime.calls]
    assert ('pool', 'drain') in [(r, s) for r, s, *_ in runtime.calls]


def test_appkit_partial_lookup_failure_keeps_bundle_executable_and_pid_for_recovery():
    runtime = CocoaRuntime()
    launcher = native_launcher(runtime)
    launcher._lookup = type('Lookup', (), {'snapshot': lambda self, pid, **kwargs: None})()
    with pytest.raises(api().ExcelProcessError, match='EXCEL_LAUNCH_IDENTITY_UNAVAILABLE'):
        launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    assert launcher.recovery_candidate['pid'] == 4242
    assert launcher.recovery_candidate['bundle_id'] == 'com.microsoft.Excel'
    assert launcher.recovery_candidate['executable'] == identity().executable
    assert launcher.recovery_candidate['launch_date'] == 100.00025


def test_nonzero_helper_reply_quarantines_even_when_it_contains_json(tmp_path):
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    script = tmp_path / 'mutation.applescript'
    owner.reserve_workbook(Path('/tmp/planned.xlsx'))
    script.write_text('mutation source')
    transport.run = lambda *args: subprocess.CompletedProcess([], 1, '{}', 'helper error')
    with pytest.raises(OSATransportError, match='EXCEL_COMMAND_FAILED'):
        owner.run(script, deadline=5)
    assert owner.state == 'quarantined'


def test_process_drift_during_readiness_reply_never_becomes_owned():
    owner, launcher, transport, _ = setup_owner()
    original_run = transport.run
    def run(*args):
        result = original_run(*args)
        launcher.current = replace(identity(), start_seconds=200)
        return result
    transport.run = run
    with pytest.raises(OSATransportError, match='OSA_PROCESS_IDENTITY_CHANGED'):
        owner.start(deadline=5)
    assert owner.state == 'quarantined'
    assert not owner.is_private_owned


def test_launch_date_is_independent_observation_not_libproc_birth_equality():
    owner, launcher, _, _ = setup_owner()
    original = launcher.launch
    def launch(*args):
        launcher.record = replace(original(*args), launch_date=99.985296)
        return launcher.record
    launcher.launch = launch
    owner.start(deadline=5)
    assert owner.is_private_owned
    assert owner.identity.start_microseconds == 250
    assert owner.launch_date == 99.985296


def test_claim_without_reservation_cannot_adopt_existing_workbook():
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    transport.books = [book()]
    with pytest.raises(ValueError):
        owner.claim_workbook(Path('/tmp/owned.xlsx'), deadline=5)
    assert transport.events == ['inventory']
    assert owner.owned_path is None


def test_ready_run_without_reservation_never_dispatches(tmp_path):
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    script = tmp_path / 'mutation.applescript'
    script.write_text('-- owner operation: mutate')
    with pytest.raises(api().ExcelProcessError, match='EXCEL_WORKBOOK_NOT_RESERVED'):
        owner.run(script, deadline=5)
    assert transport.events == ['inventory']


def test_claim_requires_saved_workbook_after_reservation():
    owner, _, transport, _ = setup_owner()
    owner.start(deadline=5)
    owner.reserve_workbook(Path('/tmp/owned.xlsx'))
    transport.books = [dict(book(), saved=False)]
    with pytest.raises(api().ExcelProcessError, match='EXCEL_WORKBOOK_NOT_SAVED'):
        owner.claim_workbook(Path('/tmp/owned.xlsx'), deadline=5)
    assert owner.state == 'quarantined'
    assert owner.recovery['owned_path'] == '/tmp/owned.xlsx'


def test_native_command_failure_retains_evidence_in_requested_directory(tmp_path):
    evidence = tmp_path / 'owner'
    owner, _, transport, _ = setup_owner(evidence_dir=evidence)
    transport.failure = OSATransportError('OSA_TIMEOUT', stdout='partial output',
        stderr='native timeout', outcome_uncertain=True)
    with pytest.raises(OSATransportError):
        owner.start(deadline=5)
    assert owner.recovery['evidence_dir'] == str(evidence)
    command = owner.recovery['last_command']
    assert Path(command['script_path']).is_file()
    assert Path(command['stdout_path']).read_text() == 'partial output'
    assert Path(command['stderr_path']).read_text() == 'native timeout'
    assert command['error']['code'] == 'OSA_TIMEOUT'
    assert command['stdout'] == 'partial output'
    assert json.loads(Path(command['diagnostic_path']).read_text())['error']['code'] == 'OSA_TIMEOUT'


def test_successful_internal_commands_remain_in_private_evidence_directory():
    owner, _, _, _ = setup_owner()
    owner.start(deadline=5)
    assert owner.evidence_dir.is_dir()
    assert owner.last_command['error'] is None
    assert Path(owner.last_command['script_path']).is_file()
    assert json.loads(Path(owner.last_command['stdout_path']).read_text())['version'] == '16.101'


def test_bad_ack_is_recorded_with_raw_reply_and_retained_script(tmp_path):
    owner, _, transport, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    transport.bad_nonce = True
    with pytest.raises(api().ExcelProcessError, match='EXCEL_BAD_ACK'):
        owner.start(deadline=5)
    command = owner.recovery['last_command']
    assert command['error']['code'] == 'EXCEL_BAD_ACK'
    assert json.loads(command['stdout'])['nonce'] == 'wrong'
    assert Path(command['script_path']).exists()


def test_inventory_materializes_each_workbook_by_ordinal_before_reading_properties():
    # Excel returns -50 for properties of the implicit collection reference from
    # `repeat with b in workbooks`; generated AppleScript must use an indexed
    # workbook specifier. Native run01 is the execution evidence for this guard.
    source = api()._script('inventory', 'a' * 32, None)
    loop = re.search(r'repeat with (\w+) from 1 to \(count of workbooks\)\n(.*?)\nend repeat', source, re.S)
    assert loop is not None, 'Inventory must enumerate workbook ordinals'
    index, body = loop.groups()
    assert body.lstrip().startswith('set b to workbook ' + index + '\n')
    assert body.index('set b to workbook ' + index) < body.index('my j(name of b)')


@pytest.mark.parametrize('failed_file', ['stdout.txt', 'stderr.txt', 'diagnostic.json'])
def test_diagnostic_failure_does_not_mask_native_error_and_attempts_every_file(tmp_path, monkeypatch, failed_file):
    owner, _, transport, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    native_error = OSATransportError('OSA_TIMEOUT', stdout='partial', stderr='native failure', outcome_uncertain=True)
    transport.failure = native_error
    original = Path.write_text
    attempted = []
    def write(path, *args, **kwargs):
        if path.name in ('stdout.txt', 'stderr.txt', 'diagnostic.json'):
            attempted.append(path.name)
            if path.name == failed_file:
                raise OSError('evidence disk unavailable')
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(OSATransportError) as caught:
        owner.start(deadline=5)
    assert caught.value is native_error
    assert attempted == ['stdout.txt', 'stderr.txt', 'diagnostic.json']
    assert owner.state == 'quarantined'
    assert owner.recovery['code'] == 'OSA_TIMEOUT'
    command = owner.recovery['last_command']
    assert command is owner.last_command
    assert command['stdout'] == 'partial'
    assert command['stderr'] == 'native failure'
    assert command['error']['code'] == 'OSA_TIMEOUT'
    assert command['diagnostic_errors'][0]['path'].endswith(failed_file)


def test_all_diagnostic_writes_fail_after_native_success_quarantines_with_reply_in_memory(tmp_path, monkeypatch):
    owner, _, transport, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    original = Path.write_text
    attempted = []
    def write(path, *args, **kwargs):
        if path.name in ('stdout.txt', 'stderr.txt', 'diagnostic.json'):
            attempted.append(path.name)
            raise OSError('disk full')
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(api().ExcelProcessError, match='EXCEL_DIAGNOSTIC_WRITE_FAILED'):
        owner.start(deadline=5)
    assert attempted == ['stdout.txt', 'stderr.txt', 'diagnostic.json']
    assert transport.events == ['inventory']
    assert owner.state == 'quarantined'
    assert not owner.is_private_owned
    assert owner.recovery['code'] == 'EXCEL_DIAGNOSTIC_WRITE_FAILED'
    command = owner.recovery['last_command']
    assert json.loads(command['stdout'])['version'] == '16.101'
    assert len(command['diagnostic_errors']) == 3
    assert command['error'] is None


def test_diagnostic_cancellation_cannot_mask_existing_native_primary(tmp_path, monkeypatch):
    owner, _, transport, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    native_error = OSATransportError('OSA_TIMEOUT', stdout='partial', outcome_uncertain=True)
    transport.failure = native_error
    original = Path.write_text
    attempted = []
    cancellation = KeyboardInterrupt('cancel evidence')
    def write(path, *args, **kwargs):
        if path.name in ('stdout.txt', 'stderr.txt', 'diagnostic.json'):
            attempted.append(path.name)
            if path.name == 'stdout.txt':
                raise cancellation
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(BaseException) as caught:
        owner.start(deadline=5)
    assert caught.value is native_error
    assert attempted == ['stdout.txt', 'stderr.txt', 'diagnostic.json']
    assert owner.state == 'quarantined'
    assert owner.recovery['code'] == 'OSA_TIMEOUT'
    assert owner.last_command['diagnostic_errors'][0]['code'] == 'KeyboardInterrupt'


@pytest.mark.parametrize('cancellation, expected_code', [(KeyboardInterrupt('cancel evidence'), 'KeyboardInterrupt'), (SystemExit(7), 7)])
def test_successful_native_command_preserves_diagnostic_cancellation_after_all_writes(tmp_path, monkeypatch, cancellation, expected_code):
    owner, _, _, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    original = Path.write_text
    attempted = []
    def write(path, *args, **kwargs):
        if path.name in ('stdout.txt', 'stderr.txt', 'diagnostic.json'):
            attempted.append(path.name)
            if path.name == 'stdout.txt':
                raise OSError('first disk failure')
            if path.name == 'stderr.txt':
                raise cancellation
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(BaseException) as caught:
        owner.start(deadline=5)
    assert caught.value is cancellation
    assert attempted == ['stdout.txt', 'stderr.txt', 'diagnostic.json']
    assert owner.state == 'quarantined'
    assert owner.recovery['code'] == expected_code
    assert json.loads(owner.last_command['stdout'])['version'] == '16.101'
    assert [error['code'] for error in owner.last_command['diagnostic_errors']] == ['OSError', type(cancellation).__name__]


def test_post_command_identity_failure_preserves_successful_native_reply(tmp_path):
    owner, launcher, transport, _ = setup_owner(evidence_dir=tmp_path / 'owner')
    original = transport.run
    def run(*args):
        result = original(*args)
        launcher.current = None
        return result
    transport.run = run
    with pytest.raises(OSATransportError, match='OSA_PROCESS_NOT_FOUND'):
        owner.start(deadline=5)
    command = owner.recovery['last_command']
    assert command['returncode'] == 0
    assert json.loads(command['stdout'])['version'] == '16.101'
    assert json.loads(Path(command['stdout_path']).read_text())['workbooks'] == []
    assert command['error']['code'] == 'OSA_PROCESS_NOT_FOUND'


def test_launcher_validates_retained_launch_handle_after_pid_rediscovery_becomes_nil():
    launcher = native_launcher(CocoaRuntime())
    class Lookup:
        def __init__(self):
            self.rediscovery_available = True
            self.checked = []
        def snapshot(self, pid, *, application=None):
            self.checked.append((pid, application))
            if pid == 4242 and application == 'returned-app':
                return identity()
            return identity() if self.rediscovery_available else None
    lookup = Lookup()
    launcher._lookup = lookup
    launched = launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    lookup.rediscovery_available = False
    assert launcher.snapshot(4242) == identity()
    assert lookup.checked == [(4242, launched.handle), (4242, launched.handle)]


def test_launcher_does_not_offer_owned_handle_when_inspecting_another_pid():
    launcher = native_launcher(CocoaRuntime())
    launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    class Lookup:
        def snapshot(self, pid, *, application=None):
            assert pid == 9000
            if application is not None:
                raise AssertionError('Unrelated PID must not receive retained owner handle')
            return replace(identity(), pid=9000)
    launcher._lookup = Lookup()
    assert launcher.snapshot(9000).pid == 9000


def test_launcher_retained_handle_never_bypasses_birth_identity_revalidation():
    launcher = native_launcher(CocoaRuntime())
    launched = launcher.launch(Path('/Applications/Microsoft Excel.app'), 5)
    class Lookup:
        def snapshot(self, pid, *, application=None):
            assert application is launched.handle
            raise OSATransportError('OSA_PROCESS_IDENTITY_CHANGED', pid=pid)
    launcher._lookup = Lookup()
    with pytest.raises(OSATransportError, match='OSA_PROCESS_IDENTITY_CHANGED'):
        launcher.snapshot(4242)
