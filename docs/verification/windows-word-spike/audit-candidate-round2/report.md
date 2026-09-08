# Reviewed candidate regression audit, round 2

Code candidate: `d4884a55fe8a0db48d1f2f48543948959aa08dbf`.
Fetched review tip: `067a23c560b317e3b9f1f5d7ae65e37e01b95e47` (additional Mac
evidence only). Merge/execution HEAD: `dc6a781de0a7eb5416796d70b7952ab122f533f0`.
The independent native UI audit was committed first as `c735e9e`, then preserved
by the merge. No production or runner source was edited on this Windows host.

## Completed evidence

- Windows execution of `test_msoffice_artifact_validation.py`,
  `test_msoffice_sentinel_failures.py`, and `test_msoffice_spike_ownership.py`:
  **56 passed in 18.58 seconds**, exit 0. Exact invocation and stdout/stderr retained.
- Python 3.11.9 x64, existing PyMuPDF/Pillow. Since no installed interpreter had
  pytest, pytest 9.1.1 and its dependencies were placed only under ignored
  `build/msoffice-spike/audit-pytest-deps` using `uv pip install --target`.
  The test process prepended this directory to its own sys.path; global Python
  and application/plugin settings were not modified.
- Current validator on an independent copy of historical synthetic native run10:
  **12/12 checks passed**, exit 0, including numbered headings and actual TOC page
  numbers. This is explicitly a validator-only run on older output, not a fresh
  native Word/sentinel pass for the changed wrapper.
- The real mouse/keyboard Word edit/Undo/save/close/reopen audit already passed
  on baseline 760704f; its evidence is in `../audit-ui-round1/`. The native
  windows_word.py is unchanged in d4884a5, but this does not imply the updated
  sentinel's new failure/timeout paths have all been exercised natively.
- Exact current raw source bytes for windows_word.py, windows_unsaved_sentinel.py,
  validate_windows_artifacts.py and mac_word.py are retained with raw/normalized
  SHA256 and Git blob IDs. Normalized bytes equal candidate d4884a5.

## Fresh native sentinel gate: awaiting scope clarification

The latest instruction prohibits handling the old PID 9252. The available
GetActiveObject registration points to the legacy Word instance, so the current
wrapper would create its sentinel there. To avoid changing that instance, an
outer audit helper created a new separate, empty native Word candidate PID 14432
and tried to make it the registered task target using RegisterActiveObject.

The registration call failed before returning a cookie:

```text
pywintypes.com_error: (-2147024809, '参数错误。', None, None)
```

The original traceback is in `isolated-registration.json`; the helper source is
retained. Its overall status is `blocked`. No sentinel subprocess was launched,
no document was created by this registration experiment, no registration cookie
was obtained, and no Quit/kill/registry change was performed. The deprecation
warning for pythoncom.MakeIID is incidental; the native error is E_INVALIDARG.
The helper uses COM's running-object registration API, not Windows registry edits.
The documented API signature was checked against the
[pywin32 documentation](https://mhammond.github.io/pywin32/pythoncom__RegisterActiveObject_meth.html).

A scope question was sent: whether the user permits creating and closing only a
task-owned unsaved sentinel in the already registered legacy instance, without
quitting it or changing existing documents. Until answered, no mutation of that
instance is authorized by this audit. Therefore fresh native run11 / sentinel04
and its current-validator check remain **unrun**, and no new native acceptance
is claimed. Production admission remains **NO_GO**.

## Coordination and publication notes

An attempted status message to controller task
`01a07ba7-13a8-7a83-8a60-821531cf8261` (host local) returned
`no rollout found for thread id ...`; progress was reported in the current task.
An interim ordinary Git push returned `Recv failure: Connection was reset`.
Both the UI audit commit and candidate merge remained local at that point.
This evidence commit preserves completed work and the exact unresolved gate so
the controller can review it without mistaking older native successes for this
new candidate's native execution.
