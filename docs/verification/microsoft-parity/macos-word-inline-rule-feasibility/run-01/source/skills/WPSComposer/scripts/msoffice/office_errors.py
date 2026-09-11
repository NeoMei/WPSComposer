"""Platform-independent, closed diagnostics for Microsoft Office adapters."""
from __future__ import annotations

_MESSAGES = {
    'NATIVE_OFFICE_UNAVAILABLE': 'The requested native Microsoft application is unavailable.',
    'NATIVE_OFFICE_TIMEOUT': 'Native Office timed out; uncertain task files are retained.',
    'NATIVE_OFFICE_QUARANTINED': 'Native Office cleanup is unverified; explicit recovery is required.',
    'NATIVE_OFFICE_EXECUTION_FAILED': 'Native Office did not produce a verified artifact; private diagnostics are retained.',
    'NATIVE_OFFICE_UNSUPPORTED': 'The native Office operation is not supported.',
}
NATIVE_OFFICE_ERROR_CODES = frozenset(_MESSAGES)


class NativeOfficeError(RuntimeError):
    def __init__(self, code, *, staging_path=None, diagnostic_path=None, quarantine_path=None):
        self.code = code
        self.safe_message = _MESSAGES[code]
        for name, value in (('staging_path', staging_path), ('diagnostic_path', diagnostic_path),
                            ('quarantine_path', quarantine_path)):
            setattr(self, name, str(value) if value is not None else None)
        super().__init__(self.safe_message)
