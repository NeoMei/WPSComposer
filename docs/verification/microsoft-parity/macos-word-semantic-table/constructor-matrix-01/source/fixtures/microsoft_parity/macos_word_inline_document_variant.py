"""Compile-only constructor hypothesis; no native runner or import side effects.

The sole constructor variant changes the creation container/range argument.
Read-only diagnostics precede the unchanged count/type/geometry gate so a failed
future native run retains enough information to distinguish collection behavior.
Fresh explicit native lease and canonical owned-session/sentinel runner required.
"""
from __future__ import annotations

from fixtures.microsoft_parity.macos_word_rules_probe import build_probe_commands


def build_commands():
    commands = build_probe_commands('inline')
    index = commands.index('make new standard inline horizontal line at insertionRange')
    commands[index:index + 1] = [
        'set creationResult to make new standard inline horizontal line at boundDoc with properties {text object:insertionRange}',
        'set totalInlineAfter to count inline shapes of boundDoc',
        'set ownedCharacterCount to count characters of text object of boundDoc',
        # A subtype collection failure is evidence, not a fallback success.
        'set standardCountAfter to -1',
        'set standardCountError to ""',
        'try',
        'set standardCountAfter to count standard inline horizontal lines of boundDoc',
        'on error diagnosticMessage number diagnosticNumber',
        'set standardCountError to (diagnosticNumber as text) & ":" & diagnosticMessage',
        'end try',
        'log my jsonRows({{"inline-creation-counts",beforeCount,totalInlineAfter,standardCountAfter,ownedCharacterCount,standardCountError}})',
        # AppleScript's log formatter retains the returned native class without
        # trying to coerce Word's application class constant into a JSON string.
        'try',
        'log {"inline-creation-result-type",class of creationResult}',
        'on error diagnosticMessage number diagnosticNumber',
        'log {"inline-creation-result-type-error",diagnosticNumber,diagnosticMessage}',
        'end try',
    ]
    return commands
