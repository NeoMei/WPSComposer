"""Compare manifest hashes with worktree files and exact Git blobs."""
import hashlib, json, subprocess, sys
from pathlib import Path
root = Path('docs/verification/msoffice-production/windows')
entries = []
for line in (root/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines():
    expected, relative = line.split('  ', 1)
    entries.append((expected, root/relative))
requests = ''.join('HEAD:' + path.as_posix() + '\n' for _, path in entries).encode()
raw = subprocess.check_output(['git','-c','core.longpaths=true','cat-file','--batch'], input=requests)
offset = 0; failures = []
for expected, path in entries:
    end = raw.index(b'\n', offset); header = raw[offset:end].decode(); offset = end + 1
    if header.endswith(' missing'):
        failures.append({'path':str(path), 'error':'missing Git blob'}); continue
    oid, kind, size = header.split(); size = int(size)
    blob = raw[offset:offset+size]; offset += size + 1
    actual_git = hashlib.sha256(blob).hexdigest()
    actual_disk = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual_git != expected or actual_disk != expected:
        failures.append({'path':str(path),'expected':expected,'git':actual_git,'disk':actual_disk})
result = {'head':subprocess.check_output(['git','rev-parse','HEAD']).decode().strip(), 'entries':len(entries), 'failures':failures, 'status':'PASS' if not failures else 'FAIL'}
Path(sys.argv[1]).write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result)); raise SystemExit(bool(failures))
