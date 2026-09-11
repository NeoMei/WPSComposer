# Windows parity candidate handoff

This is a new three-application candidate. The previous v0.9.0 Word acceptance
does not validate these changes. The [September10 checkpoint](windows-70d3d86-20260910/SUMMARY.md) records blocked startup on70d3d86 and older native failures, not a native pass. Recover the exact retained objects before any new Office run; the existing Excel instance may still contain an unsaved workbook. A completed remote task is not a completed acceptance matrix.

Use a separate checkout of the reviewed `codex/microsoft-parity` candidate and
record its exact commit with `git rev-parse HEAD`. Preserve the existing Windows
checkout and all user documents. Use the existing task `确认远程目录写入权限` for
coordination; do not mistake a successful message dispatch for actual execution.

In that isolated checkout, create a private Python environment if needed:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e '.[dev,windows]'
# Node.js22 and npm are also needed for the pinned fixture templates.
node --version
npm ci --prefix macos/wps-jsapi-probe --ignore-scripts
.\.venv\Scripts\python.exe -m pytest tests/macos_probe/test_profile_server.py -q
.\.venv\Scripts\python.exe -m pytest -q
```

Do not skip the npm setup: the older isolated Windows checkout lacked
`wpsjs` template resources, producing setup failures during otherwise portable
tests. Preserve the old raw log; a later successful setup does not rewrite it.

Confirm Word, Excel and PowerPoint are installed and available in the interactive
desktop. Run the guarded standard acceptance with a **new** output directory:

```powershell
.\.venv\Scripts\python.exe fixtures\microsoft_parity\windows_native.py --checkout . --output docs\verification\microsoft-parity\windows-native-01
```

Use `--kind writer`, `--kind sheet` or `--kind slide` to isolate a failed app in a
new output directory after its exact recovery is complete. A failed or timed-out
case stops the runner before later applications. Never rerun blindly over a
quarantine or delete recovery evidence to make a new run start.

The runner checks public creation, generation, conversion, inspection, editing,
typed structural handles, exact reopened data and PDF content. It records saved
and unsaved test sentinels, original document identities and content hashes,
native executable/PID/window identity, proxy close acknowledgements and Python
worker exit. It must preserve unrelated documents and processes. Importing the
runner or running its portable tests does not execute Office.

The optional `--fault-timeout` phase deliberately retains uncertain owned Office
documents and component recovery files. Run it only after normal acceptance and
with the exact retained-document recovery procedure available. Its timeout may
terminate the precise Python child; it must never kill Office or close unknown
documents. A retained file is not proof of successful cleanup.

After standard native acceptance, separately exercise actual Word/Excel/
PowerPoint UI edit → Undo → explicit save → close → reopen on test outputs.
Inspect the final layout and formulas/shapes, and verify unrelated sentinels.
Repeat the public smoke from an isolated installed bundle, leaving the personal
marketplace unchanged. Preserve raw reports, failed runs and corrected reruns.

Every accepted capability claim needs its own report-declared check mapping,
platform/component/engine, exact source digest and report hash. The evidence gate
does not convert generic `passed` reports into full parity. Full catalog,
cross-engine WPS regression, installation and UI gates remain independent.
