from __future__ import annotations

import copy
import hashlib
import json

import pytest

from fixtures.microsoft_parity.evidence_gate import evaluate, source_digest


@pytest.fixture
def case(tmp_path):
    (tmp_path / 'skills').mkdir()
    (tmp_path / 'skills' / 'engine.py').write_text('version = 1\n')
    records = [{'id': 'composer.writer.save', 'required': True, 'component': 'writer'}]
    claims = []
    for platform in ('darwin', 'win32'):
        evidence = tmp_path / (platform + '.json')
        evidence.write_text(json.dumps({'passed': True, 'platform': platform, 'engine': 'msoffice',
                                       'component': 'writer',
                                       'capability_checks': {'composer.writer.save': ['reopen']},
                                       'source_digest': source_digest(tmp_path), 'checks': {'reopen': True}}))
        claims.append({'id': records[0]['id'], 'platform': platform,
               'implementation': 'implemented', 'verification': 'native_verified',
               'source_digest': source_digest(tmp_path),
               'evidence': [{'path': evidence.name, 'sha256': hashlib.sha256(evidence.read_bytes()).hexdigest(),
                             'checks': ['reopen']}]})
    return tmp_path, records, claims


def test_missing_claims_remain_required_and_open(case):
    root, rows, _ = case
    report = evaluate(root, rows, [])
    assert report['ready'] is False
    assert len(report['rows']) == 2
    assert all(row['state'] == 'unverified' for row in report['rows'])


def test_only_explicit_source_bound_native_checks_close_a_row(case):
    root, rows, claims = case
    report = evaluate(root, rows, claims)
    assert report['ready'] is True
    assert report['verified_count'] == 2


def test_changed_source_reopens_all_previously_verified_rows(case):
    root, rows, claims = case
    (root / 'skills' / 'engine.py').write_text('version = 2\n')
    report = evaluate(root, rows, claims)
    assert not report['ready']
    assert all('source_changed' in row['issues'] for row in report['rows'])


@pytest.mark.parametrize('change', ['bytes', 'failed', 'missing_check', 'empty_checks'])
def test_evidence_cannot_pass_on_path_or_report_existence_alone(case, change):
    root, rows, claims = case
    path = root / 'darwin.json'
    if change == 'bytes':
        path.write_text(path.read_text() + ' ')
    elif change == 'empty_checks':
        claims[0]['evidence'][0]['checks'] = []
    else:
        data = json.loads(path.read_text())
        if change == 'failed':
            data['passed'] = False
        else:
            data['checks'] = {}
        path.write_text(json.dumps(data))
        claims[0]['evidence'][0]['sha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
    assert not evaluate(root, rows, claims)['ready']


@pytest.mark.parametrize('change', ['duplicate', 'unknown', 'platform', 'traversal', 'absolute'])
def test_invalid_claim_identity_or_evidence_path_is_rejected(case, change):
    root, rows, claims = case
    if change == 'duplicate':
        claims.append(copy.deepcopy(claims[0]))
    elif change == 'unknown':
        claims[0]['id'] = 'imaginary'
    elif change == 'platform':
        claims[0]['platform'] = 'linux'
    else:
        claims[0]['evidence'][0]['path'] = '../native.json' if change == 'traversal' else str(root / 'native.json')
    with pytest.raises(ValueError):
        evaluate(root, rows, claims)


def test_unit_tests_and_partial_implementation_do_not_close_native_gate(case):
    root, rows, claims = case
    claims[0]['verification'] = 'unit_verified'
    claims[1]['implementation'] = 'partial'
    report = evaluate(root, rows, claims)
    assert not report['ready'] and report['verified_count'] == 0


@pytest.mark.parametrize('field,value', [('platform', 'win32'), ('engine', 'wps'), ('source_digest', 'old')])
def test_report_itself_must_bind_engine_platform_and_source(case, field, value):
    root, rows, claims = case
    path = root / 'darwin.json'
    report = json.loads(path.read_text())
    report[field] = value
    path.write_text(json.dumps(report))
    claims[0]['evidence'][0]['sha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
    assert not evaluate(root, rows, claims)['ready']


def test_evidence_directory_and_python_cache_do_not_change_source_digest(case):
    root, _, _ = case
    before = source_digest(root)
    (root / 'skills' / '__pycache__').mkdir()
    (root / 'skills' / '__pycache__' / 'engine.pyc').write_bytes(b'cache')
    (root / 'native.json').write_text('changed report')
    assert source_digest(root) == before


@pytest.mark.parametrize('name', ['package.json', 'package-lock.json'])
def test_shipped_runtime_dependencies_are_bound_to_source_digest(case, name):
    root, _, _ = case
    package = root / 'macos' / 'wps-jsapi-probe' / name
    package.parent.mkdir(parents=True)
    package.write_text('{"version": 1}')
    before = source_digest(root)
    package.write_text('{"version": 2}')
    assert source_digest(root) != before


def _rewrite_report(root, claims, mutate):
    path = root / 'darwin.json'
    report = json.loads(path.read_text())
    mutate(report)
    path.write_text(json.dumps(report))
    claims[0]['evidence'][0]['sha256'] = hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.mark.parametrize('change', ['component', 'missing_component', 'different_capability',
                                  'missing_mapping', 'empty_mapping', 'wrong_check', 'unknown_mapping'])
def test_generic_success_cannot_certify_an_unrelated_capability(case, change):
    root, rows, claims = case
    def mutate(report):
        if change == 'component':
            report['component'] = 'presentation'
        elif change == 'missing_component':
            del report['component']
        elif change == 'different_capability':
            report['capability_checks'] = {'public.presentation.generate': ['reopen']}
        elif change == 'missing_mapping':
            del report['capability_checks']
        elif change == 'empty_mapping':
            report['capability_checks'] = {'composer.writer.save': []}
        elif change == 'wrong_check':
            report['checks']['generate'] = True
            report['capability_checks'] = {'composer.writer.save': ['generate']}
        else:
            report['capability_checks']['invented.capability'] = ['reopen']
    _rewrite_report(root, claims, mutate)
    assert not evaluate(root, rows, claims)['ready']


def test_claim_cannot_omit_a_failed_mapped_check(case):
    root, rows, claims = case
    def mutate(report):
        report['capability_checks']['composer.writer.save'].append('source_preserved')
        report['checks']['source_preserved'] = False
    _rewrite_report(root, claims, mutate)
    assert not evaluate(root, rows, claims)['ready']


def test_common_capability_requires_explicit_common_component(case):
    root, rows, claims = case
    rows[0].update(id='composer.common.close', component='common')
    for claim in claims:
        claim['id'] = 'composer.common.close'
        path = root / claim['evidence'][0]['path']
        report = json.loads(path.read_text())
        report.update(component='common', capability_checks={'composer.common.close': ['reopen']})
        path.write_text(json.dumps(report))
        claim['evidence'][0]['sha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
    assert evaluate(root, rows, claims)['ready']
    _rewrite_report(root, claims, lambda report: report.update(component='presentation'))
    assert not evaluate(root, rows, claims)['ready']


def test_unshipped_runner_and_test_sources_do_not_change_installed_source_digest(case):
    root, _, _ = case
    before = source_digest(root)
    for directory in ('fixtures/microsoft_parity', 'tests/msoffice'):
        path = root / directory / 'runner.py'
        path.parent.mkdir(parents=True)
        path.write_text('changed verification code')
    assert source_digest(root) == before


@pytest.mark.parametrize('names', [None, 'reopen', {}, [None], [''], ['reopen', 'reopen'], [['reopen']]])
def test_malformed_capability_check_mapping_fails_closed(case, names):
    root, rows, claims = case
    _rewrite_report(root, claims, lambda report: report.update(
        capability_checks={'composer.writer.save': names}))
    assert not evaluate(root, rows, claims)['ready']


def test_report_cannot_mix_components_under_one_component_declaration(case):
    root, rows, claims = case
    rows.append({'id': 'public.presentation.generate', 'required': False, 'component': 'presentation'})
    def mutate(report):
        report['capability_checks']['public.presentation.generate'] = ['reopen']
    _rewrite_report(root, claims, mutate)
    result = evaluate(root, rows, claims)
    assert not result['ready']
    assert 'capability_binding_mismatch:darwin.json' in result['rows'][0]['issues']
