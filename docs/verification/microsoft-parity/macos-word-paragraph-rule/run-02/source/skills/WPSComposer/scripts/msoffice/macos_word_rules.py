"""Native paragraph-border rule, bound to the session's terminal body range."""
from __future__ import annotations

import json
from uuid import uuid4

from .errors import NativeWordError


def paragraph_rule_commands(session):
    """Rule the current terminal paragraph, then clear the new following border."""
    return session._position('end') + [
        'set ruleOriginalPoint to insertionPoint',
        'set ruleBeforeEnd to end of content of text object of boundDoc',
        'set ruleBeforeParagraphs to count paragraphs of boundDoc',
        'set rulePrefix to ""',
        'if insertionPoint > 0 then',
        'set rulePrefixRange to create range boundDoc start 0 end insertionPoint',
        'set rulePrefix to content of rulePrefixRange as text',
        'end if',
    ] + [
        'set ruleRange to text object of paragraph ruleBeforeParagraphs of boundDoc',
        'set ruleStart to start of content of ruleRange',
        'set alignment of paragraph format of ruleRange to align paragraph center',
        'set ownBorder to get border ruleRange which border border bottom',
        'set line style of ownBorder to line style single',
        'set line width of ownBorder to line width75 point',
        'set color of ownBorder to {49344,49344,49344}',
        'set insertionRange to create range boundDoc start ruleOriginalPoint end ruleOriginalPoint',
        'set content of insertionRange to " " & return',
        'set ruleRange to text object of paragraph ruleBeforeParagraphs of boundDoc',
        'set ownBorder to get border ruleRange which border border bottom',
        'set insertedRange to create range boundDoc start ruleOriginalPoint end (ruleOriginalPoint + 2)',
        'set followingRange to create range boundDoc start (ruleOriginalPoint + 2) end (ruleOriginalPoint + 3)',
        # Clear only the inherited border. Preserve following style/spacing and
        # require the centered alignment inherited from the ruled paragraph.
        'set followingStyle to name local of style of followingRange as text',
        'set followingFormat to paragraph format of followingRange',
        'set followingBefore to space before of followingFormat',
        'set followingAfter to space after of followingFormat',
        'set followingLeft to paragraph format left indent of followingFormat',
        'set followingRight to paragraph format right indent of followingFormat',
        'set followingFirst to first line indent of followingFormat',
        'set followingSpacing to line spacing of followingFormat',
        'set followingSpacingRule to line spacing rule of followingFormat',
        'set followingBorder to get border followingRange which border border bottom',
        'set line style of followingBorder to line style none',
        'set followingFormat to paragraph format of followingRange',
        'set followingPreserved to ((name local of style of followingRange as text) is followingStyle '
        'and space before of followingFormat is followingBefore and space after of followingFormat is followingAfter '
        'and paragraph format left indent of followingFormat is followingLeft '
        'and paragraph format right indent of followingFormat is followingRight '
        'and first line indent of followingFormat is followingFirst '
        'and line spacing of followingFormat is followingSpacing '
        'and line spacing rule of followingFormat is followingSpacingRule '
        'and (alignment of followingFormat is align paragraph center))',
        'set prefixPreserved to true',
        'if ruleOriginalPoint > 0 then',
        'set rulePrefixRange to create range boundDoc start 0 end ruleOriginalPoint',
        "set prefixPreserved to ((current application's NSString's stringWithString:(content of rulePrefixRange as text))'s isEqualToString:rulePrefix) as boolean",
        'end if',
        'set nativeRows to {{"paragraph-rule",ruleOriginalPoint,ruleStart,'
        'end of content of ruleRange,ruleBeforeEnd,end of content of text object of boundDoc,'
        'ruleBeforeParagraphs,count paragraphs of boundDoc,prefixPreserved,'
        'content of insertedRange as text,content of followingRange as text,'
        '(alignment of paragraph format of ruleRange is align paragraph center),'
        '(line style of ownBorder is line style single),'
        '(line width of ownBorder is line width75 point),'
        '(color of ownBorder is {49344,49344,49344}),'
        '(line style of followingBorder is line style none),followingPreserved}}',
    ]


def _valid_ack(rows):
    if not isinstance(rows, list) or len(rows) != 1:
        return False
    row = rows[0]
    if not (isinstance(row, list) and len(row) == 17
            and row[0] == 'paragraph-rule'
            and all(type(v) is int for v in row[1:8])):
        return False
    original, start, end, before_end, after_end, before_paras, after_paras = row[1:8]
    return (0 <= start <= original and (original > 0 or start == 0)
            and end == original + 2 and before_end == original + 1 and after_end == end + 1
            and before_paras >= 1 and after_paras == before_paras + 1
            and row[9:11] == [' \r', '\r']
            and all(row[i] is True for i in (8, 11, 12, 13, 14, 15, 16)))


def add_paragraph_horizontal_line(session):
    """Insert a centered silver .75pt bottom rule and a clean trailing paragraph."""
    session._mutation_preflight()
    try:
        rows = session._execute_topology_mutation(paragraph_rule_commands(session))
    except NativeWordError:
        # Even a definite native command error can follow a successful prefix.
        # Keep the original typed error and private diagnostic; never retry.
        session._retain_evidence = True
        session._retain('Paragraph rule native completion unverified')
        raise
    if not _valid_ack(rows):
        session._retain_evidence = True
        session._retain('Paragraph rule acknowledgement invalid')
        diagnostic = session.staging_root / (uuid4().hex + '-paragraph-rule-ack.json')
        diagnostic.write_text(json.dumps(rows, ensure_ascii=False, default=repr), encoding='utf-8')
        diagnostic.chmod(0o600)
        raise NativeWordError('NATIVE_WORD_QUARANTINED', staging_path=session.staging_root,
                              diagnostic_path=diagnostic,
                              quarantine_path=session.lock.quarantine_path if session.lock else None)
    session._structural_changed = True
    session._pending_heading = None
