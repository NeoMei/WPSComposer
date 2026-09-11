"""Offline checker correction only: never call native Office or alter run-02."""
from pathlib import Path
import hashlib
import json
import shutil

from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture

root = Path.cwd()
original = root/'docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02'
output = original.parent/'run-02-offline-correction-02'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inventory(directory):
    return {str(path.relative_to(directory)):sha(path) for path in sorted(directory.rglob('*')) if path.is_file()}

before = inventory(original)
prior = json.loads((original/'report.json').read_text())
assert prior['status'] == 'FAIL'
assert [key for key, value in prior['checks'].items() if not value] == ['pdf_rule_drawn']
for name, expected in prior['artifact_hashes'].items():
    assert sha(original/name) == expected, name
output.mkdir(exist_ok=False)
for name in ('paragraph-rule.docx', 'paragraph-rule.pdf'):
    shutil.copyfile(original/name, output/name)
source_hashes = fixture.retain_sources(output, fixture.SOURCES)
script_copy = output/'offline-revalidate.py'
shutil.copyfile(__file__, script_copy)
corrected_artifact_checks = fixture.artifacts(output)
assert all(corrected_artifact_checks.values())
retained_ack_checks = {'retained_native_ack_valid': fixture.native_valid(prior['native_rows']),
                       'retained_reopen_ack_valid': fixture.native_valid(prior['reopened_rows'])}
assert all(retained_ack_checks.values())
after = inventory(original)
assert before == after
report = {
    'status':'OFFLINE_CORRECTED_PASS',
    'scope':'Read-only revalidation of immutable run-02 artifacts and retained ACK rows; no new native execution',
    'original_report':str((original/'report.json').relative_to(root)),
    'original_report_sha256':before['report.json'], 'original_status':prior['status'],
    'original_checks':prior['checks'], 'corrected_artifact_checks':corrected_artifact_checks,
    'retained_ack_checks':retained_ack_checks,
    'original_tree_unchanged':before == after, 'original_file_hashes':before,
    'checker_source_hashes':source_hashes, 'offline_script_sha256':sha(script_copy),
    'revalidated_artifact_hashes':{name:sha(output/name) for name in ('paragraph-rule.docx', 'paragraph-rule.pdf', 'document.xml', 'styles.xml')},
}
(output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
print(json.dumps({'status':report['status'], 'corrected_artifact_checks':corrected_artifact_checks,
                  'retained_ack_checks':retained_ack_checks, 'original_tree_unchanged':True,
                  'report':str(output/'report.json')}, ensure_ascii=False))
