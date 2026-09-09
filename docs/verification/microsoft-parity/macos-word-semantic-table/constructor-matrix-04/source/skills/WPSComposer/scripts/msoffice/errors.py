"""Closed public diagnostics for native Word, separate from private engine logs."""
from __future__ import annotations

from pathlib import Path


_MESSAGES = {
    'NATIVE_WORD_FIELD_IDENTITY_STALE': 'Native Word field identity is stale; reopen the document to rebuild semantic tracking.',
    'NATIVE_WORD_UNSUPPORTED': 'Native Word does not support this generation capability.',
    'NATIVE_WORD_TIMEOUT': 'Native Word timed out; uncertain task files are retained.',
    'NATIVE_WORD_QUARANTINED': 'Native Word cleanup is unverified; task files are retained and recovery is required.',
    'NATIVE_WORD_EXECUTION_FAILED': 'Native Word execution failed; private evidence is retained.',
    'NATIVE_WORD_UNAVAILABLE': 'Native Microsoft Word is unavailable or not initialized.',
}
NATIVE_WORD_ERROR_CODES = frozenset(_MESSAGES)
RECOVERY_FIELDS = ('staging_path', 'diagnostic_path', 'quarantine_path')


class NativeWordError(RuntimeError):
    """Only fixed messages and task-generated locations may cross public APIs."""
    def __init__(self, code, *, staging_path=None, diagnostic_path=None, quarantine_path=None):
        if code not in _MESSAGES:
            raise ValueError('Unknown native Word error code')
        self.code = code
        for name, value in zip(RECOVERY_FIELDS, (staging_path, diagnostic_path, quarantine_path)):
            if value is not None:
                value = str(Path(value).absolute())
                if any(ord(c) < 32 for c in value):
                    raise ValueError('Invalid native Word recovery location')
            setattr(self, name, value)
        self.safe_message = _MESSAGES[code]
        for name in RECOVERY_FIELDS:
            value = getattr(self, name)
            if value is not None:
                self.safe_message += f' {name}: {value}'
        super().__init__(self.safe_message)


class NativeWordTimeoutError(NativeWordError, TimeoutError):
    def __init__(self, **locations):
        super().__init__('NATIVE_WORD_TIMEOUT', **locations)


class NativeWordCapabilityError(NativeWordError, ValueError):
    """Keep compiler exception compatibility, expose only a closed feature label."""
    def __init__(self, detail):
        super().__init__('NATIVE_WORD_UNSUPPORTED')
        labels = (
            ('equation', 'native equations'), ('color', 'custom font colors'),
            ('table', 'this table layout or cell semantics'), ('figure', 'this figure layout'),
            ('resource', 'this image resource type'), ('number', 'this numbering configuration'),
            ('footer', 'this footer configuration'), ('style', 'these style attributes'),
            ('page', 'these page settings'), ('paragraph', 'these paragraph attributes'),
        )
        for token, label in labels:
            if token in str(detail).lower():
                self.safe_message = f'Native Word does not support {label} on macOS.'
                break
        # Direct compiler callers historically receive a ValueError with detail.
        # The lifecycle and conversion boundary use safe_message, never args.
        self.args = (str(detail),)
