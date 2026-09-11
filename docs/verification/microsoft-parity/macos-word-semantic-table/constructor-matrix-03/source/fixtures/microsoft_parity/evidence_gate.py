"""Audit explicit parity claims against the frozen inventory and exact sources.

Existing representative reports are not retroactively certified. A native runner
must capture the source digest when it runs and name its actual passing checks.
Reports declare one ``component`` (including explicit ``common``) and a
``capability_checks`` object mapping frozen capability IDs to nonempty lists
of named checks. Each evidence item's checks must exactly cover its report's
mapping for that capability; a generic passing check cannot certify another row.
All mapped IDs must belong to the declared component and the frozen inventory.

The source digest binds shipped runtime code, manifests, dependency lockfiles,
and installers. Repository-only fixture runners and tests are not installed
runtime inputs: their review/provenance is separate from the installed source
digest, so this gate does not claim runner integrity from that digest.
This gate verifies evidence integrity, not the truth of an arbitrary assertion;
runner review and native/UI acceptance remain required.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from skills.WPSComposer.scripts.capability_catalog import BASELINE_COMMIT, capability_records


def source_digest(root):
    """Hash shipped source paths and bytes; evidence and caches are excluded."""
    root = Path(root).resolve(strict=True)
    files = []
    for directory in ('skills', '.codex-plugin', 'macos/wps-jsapi-probe/addin'):
        base = root / directory
        if base.exists():
            files.extend(path for path in base.rglob('*') if path.is_file()
                         and not {'__pycache__', 'node_modules'}.intersection(path.parts)
                         and path.suffix not in ('.pyc', '.pyo'))
    for name in ('install.py', 'install.sh', 'install.ps1', 'pyproject.toml',
                 'macos/wps-jsapi-probe/package.json',
                 'macos/wps-jsapi-probe/package-lock.json'):
        if (root / name).is_file():
            files.append(root / name)
    digest = hashlib.sha256()
    for path in sorted(set(files)):
        digest.update(path.relative_to(root).as_posix().encode('utf-8') + b'\0')
        digest.update(hashlib.sha256(path.read_bytes()).digest())
    return digest.hexdigest()


def _evidence_path(root, value):
    if not isinstance(value, str) or not value:
        raise ValueError('Evidence path must be a nonempty repository-relative path')
    relative = Path(value)
    if relative.is_absolute() or '..' in relative.parts:
        raise ValueError('Evidence path escapes repository')
    path = (root / relative).resolve()
    if not path.is_relative_to(root):
        raise ValueError('Evidence path escapes repository')
    return path


def _check_names(value):
    return (isinstance(value, list) and bool(value)
            and all(isinstance(name, str) and bool(name) for name in value)
            and len(set(value)) == len(value))


def _capability_binding(report, record, baseline, checks):
    if not isinstance(report, dict) or report.get('component') != record['component']:
        return False
    mapping = report.get('capability_checks')
    if not isinstance(mapping, dict) or record['id'] not in mapping:
        return False
    for identifier, names in mapping.items():
        if (identifier not in baseline or not _check_names(names)
                or baseline[identifier]['component'] != report['component']):
            return False
    return _check_names(checks) and set(checks) == set(mapping[record['id']])


def evaluate(root, records, claims):
    root = Path(root).resolve(strict=True)
    baseline = {row['id']: row for row in records}
    if len(baseline) != len(records) or not baseline:
        raise ValueError('Baseline must be nonempty and contain unique IDs')
    indexed = {}
    for claim in claims:
        key = (claim['id'], claim['platform'])
        if key[0] not in baseline or key[1] not in ('darwin', 'win32') or key in indexed:
            raise ValueError('Unknown or duplicate capability/platform claim')
        for item in claim.get('evidence', []):
            _evidence_path(root, item.get('path'))
        indexed[key] = claim
    digest = source_digest(root)
    rows = []
    for record in records:
        for platform in ('darwin', 'win32'):
            claim = indexed.get((record['id'], platform), {})
            issues = []
            if claim.get('implementation') != 'implemented':
                issues.append('implementation_incomplete')
            if claim.get('verification') != 'native_verified':
                issues.append('native_verification_missing')
            if claim.get('source_digest') != digest:
                issues.append('source_changed' if claim.get('source_digest') else 'source_binding_missing')
            evidence = claim.get('evidence', [])
            if not evidence:
                issues.append('evidence_missing')
            for item in evidence:
                path = _evidence_path(root, item['path'])
                try:
                    data = path.read_bytes()
                except OSError:
                    issues.append('evidence_missing:' + item['path'])
                    continue
                if hashlib.sha256(data).hexdigest() != item.get('sha256'):
                    issues.append('evidence_changed:' + item['path'])
                    continue
                try:
                    report = json.loads(data)
                except (ValueError, UnicodeError):
                    issues.append('invalid_report:' + item['path'])
                    continue
                checks = item.get('checks')
                if (not isinstance(report, dict) or report.get('platform') != platform
                        or report.get('engine') != 'msoffice' or report.get('source_digest') != digest):
                    issues.append('report_binding_mismatch:' + item['path'])
                if not _capability_binding(report, record, baseline, checks):
                    issues.append('capability_binding_mismatch:' + item['path'])
                if (not isinstance(report, dict) or report.get('passed') is not True
                        or not isinstance(checks, list) or not checks
                        or not isinstance(report.get('checks'), dict)
                        or any(not isinstance(name, str) or report['checks'].get(name) is not True for name in checks)):
                    issues.append('checks_not_proven:' + item['path'])
            rows.append({'id': record['id'], 'component': record['component'],
                         'platform': platform, 'required': record['required'],
                         'implementation': claim.get('implementation', 'unassessed'),
                         'state': 'native_verified' if not issues else 'unverified',
                         'issues': issues, 'evidence': evidence})
    required = [row for row in rows if row['required']]
    return {'baseline_commit': BASELINE_COMMIT, 'source_digest': digest,
            'ready': bool(required) and all(not row['issues'] for row in required),
            'required_count': len(required),
            'verified_count': sum(not row['issues'] for row in required), 'rows': rows}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument('--claims', type=Path, help='JSON list of explicitly reviewed capability claims')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    claims = json.loads(args.claims.read_text()) if args.claims else []
    report = evaluate(args.root, capability_records(), claims)
    with args.output.open('x', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write('\n')
    print(json.dumps({key: report[key] for key in ('ready', 'required_count', 'verified_count', 'source_digest')}))
    return 0 if report['ready'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
