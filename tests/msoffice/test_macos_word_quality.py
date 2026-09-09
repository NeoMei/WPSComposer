"""Frozen Writer differential RED contracts; no native calls or production edits.

The real Writer public methods are the semantic oracle. Only their native
bookmark/range/box boundaries are doubled. Candidate native boundaries are also
doubled, so these tests certify Python contract/state/error timing ONLY; native
ACK, rollback and all-location evidence must be tested separately.
"""
from __future__ import annotations

from copy import deepcopy
import ast
import hashlib
import importlib
import inspect
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.writer import NativeWriterObjectError, WriterComposer
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession

METHODS = ('reserve_document_quality_anchor', 'upsert_document_quality_notice',
           'add_document_quality_notice', 'add_quality_notice_at_bookmark')
ANCHOR = 'wpsc_document_quality_anchor'
ABSENT = '<absent>'


def test_frozen_writer_oracle_has_not_drifted_with_candidate():
    tree = ast.parse(inspect.getsource(WriterComposer))
    # Python 3.12 added the empty type_params AST field; it changes no behavior
    # in these unparameterized frozen methods and is normalized for Python 3.9+.
    for node in ast.walk(tree):
        if 'type_params' in getattr(node, '_fields', ()) and not getattr(node, 'type_params', []):
            node._fields = tuple(field for field in node._fields if field != 'type_params')
    names = set(METHODS) | {'_upsert_quality_notice_mapping', '_degradation_display'}
    methods = [node for node in tree.body[0].body if isinstance(node, ast.FunctionDef) and node.name in names]
    assert len(methods) == 6
    observed = hashlib.sha256('\n'.join(ast.dump(node, include_attributes=False) for node in methods).encode()).hexdigest()
    assert observed == '0f489c9e35c87878fae9c77b354c005126a02651a9e7b5b1c7c691ebef54814c', 'Frozen 6dd3a00 quality oracle changed; review baseline explicitly'


def implementation():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_quality')
    except ModuleNotFoundError as exc:
        if exc.name != 'skills.WPSComposer.scripts.msoffice.macos_word_quality':
            raise
        pytest.fail('Mac Word quality helper is not implemented; frozen differential contract remains RED')


class Poison:
    def __init__(self, events, label):
        self.events, self.label = events, label

    def __str__(self):
        self.events.append(('consume', self.label))
        raise ValueError('poison '+self.label)


class AuditedDict(dict):
    def __init__(self, events, **values):
        super().__init__(values)
        self.events = events

    def get(self, key, default=None):
        self.events.append(('get', key))
        return super().get(key, default)


class AuditedIssue:
    def __init__(self, events, **values):
        self._events, self._values = events, values

    def __getattr__(self, key):
        self._events.append(('attribute', key))
        if key not in self._values:
            raise AttributeError(key)
        return self._values[key]


class Range:
    def __init__(self, start, end, text='', paragraph_end=None):
        self.Start, self.End, self.Text = start, end, text
        self.paragraph_end = end if paragraph_end is None else paragraph_end

    def Paragraphs(self, index):
        assert index == 1
        return SimpleNamespace(Range=Range(self.Start, self.paragraph_end))


class NativeBoundary:
    """Controlled native outcomes shared by both implementations, not a model of Word."""
    def __init__(self):
        self.events = []
        self.selection_end = 9
        self.bookmarks = {'explicit': Range(20, 60, paragraph_end=31),
                          2: Range(40, 80, paragraph_end=51)}
        self.fail_reservation = False
        self.fail_box = None
        self.box_count = 0
        self.range_failure = None

    def range(self, start, end):
        if self.range_failure is not None:
            raise self.range_failure
        return Range(start, end)

    def reserve(self):
        self.events.append(('reserve', self.selection_end))
        if self.fail_reservation:
            raise RuntimeError('native reserve failed')
        self.bookmarks[ANCHOR] = Range(self.selection_end, self.selection_end)
        return self.selection_end

    def bookmark(self, name):
        self.events.append(('bookmark', name))
        if name not in self.bookmarks:
            raise KeyError(name)
        return self.bookmarks[name]

    def box(self, display, target):
        self.box_count += 1
        self.events.append(('box', target.Start, display))
        if self.fail_box == self.box_count:
            raise NativeWriterObjectError('DEGRADATION_INSERT_FAILED', 'native box failed')
        end = target.Start + len(display.encode('utf-16-le'))//2 + 2
        return SimpleNamespace(Range=Range(target.Start, end, display+'\r\x07\r\x07'))


class Bookmarks:
    def __init__(self, boundary):
        self.boundary = boundary

    def Add(self, name, target):
        assert name == ANCHOR
        assert target.Start == target.End == self.boundary.selection_end
        self.boundary.reserve()

    def Exists(self, name):
        self.boundary.events.append(('bookmark', name))
        return name in self.boundary.bookmarks

    def __call__(self, name):
        return SimpleNamespace(Range=self.boundary.bookmarks[name])

    Item = __call__


def harness(kind, monkeypatch):
    boundary = NativeBoundary()
    if kind == 'writer':
        obj = WriterComposer.__new__(WriterComposer)
        obj._app = SimpleNamespace(Selection=SimpleNamespace(End=boundary.selection_end))
        obj._doc = SimpleNamespace(Bookmarks=Bookmarks(boundary), Range=boundary.range)
        obj._insert_degradation_box = lambda display, target=None: boundary.box(display, target)
        call = lambda name, *a, **kw: getattr(obj, name)(*a, **kw)
    else:
        module = implementation()
        obj = SimpleNamespace()
        # These are the deliberately narrow native seams proposed in the brief.
        # They must be real helper boundaries, not test-only production hooks.
        monkeypatch.setattr(module, '_reserve_native', lambda session: boundary.reserve())
        monkeypatch.setattr(module, '_target_at', lambda session, position: boundary.range(position, position))

        def bookmark_target(session, bookmark_name):
            name = bookmark_name or ANCHOR
            bookmark = boundary.bookmark(name)
            position = bookmark.paragraph_end if bookmark_name else bookmark.Start
            return boundary.range(position, position)

        monkeypatch.setattr(module, '_target_bookmark', bookmark_target)
        monkeypatch.setattr(module, '_insert_box_native', lambda session, target, display: boundary.box(display, target))
        call = lambda name, *a, **kw: getattr(module, name)(obj, *a, **kw)
    return obj, boundary, call


def normalized(call):
    try:
        result = call()
        if result is None:
            return ('return', None)
        return ('return-range', result.Range.Start, result.Range.End, result.Range.Text)
    except Exception as exc:
        return ('error', type(exc).__name__, getattr(exc, 'code', None), str(exc))


def capture(obj, boundary, results):
    return {'results': results, 'events': deepcopy(boundary.events),
            'position': getattr(obj, '_quality_notice_anchor_position', ABSENT),
            'title': getattr(obj, '_quality_notice_title', ABSENT),
            'seen': getattr(obj, '_quality_notice_seen', ABSENT),
            'bookmark': ((boundary.bookmarks[ANCHOR].Start, boundary.bookmarks[ANCHOR].End)
                         if ANCHOR in boundary.bookmarks else ABSENT)}


def scenario(kind, case, monkeypatch, page=1):
    obj, native, call = harness(kind, monkeypatch)
    events = native.events
    results = []

    def do(name, *args, **kwargs):
        results.append(normalized(lambda: call(name, *args, **kwargs)))

    if case == 'repeated-empty-reserve':
        do(METHODS[0], 'First', [])
        do(METHODS[0], Poison(events, 'ignored replacement title'), None)
    elif case == 'first-title-cursor-and-dedup':
        do(METHODS[0], 'First')
        do(METHODS[1], {'code': 'NOTICE', 'message': 'one', 'nodeId': 'n1'})
        do(METHODS[1], AuditedDict(events, code='NOTICE', message=Poison(events, 'duplicate message'), nodeId='n1'))
        do(METHODS[1], {'code': 'NOTICE', 'message': 'two', 'nodeId': 'n2'})
    elif case == 'dict-object-distinction':
        do(METHODS[0], 'First')
        do(METHODS[1], AuditedDict(events, code='bad code', placement=None, nodeId='n1', fallbackText='dict fallback', message=Poison(events, 'unused dict message')))
        do(METHODS[1], AuditedIssue(events, code='bad code', placement='document', node_id='n1', fallbackText=Poison(events, 'unused object fallback'), message='object message'))
    elif case == 'empty-fallback-uses-message':
        do(METHODS[2], [AuditedDict(events, code='NOTICE', fallbackText='', message='visible')])
    elif case == 'placement-is-identity-not-layout':
        do(METHODS[0], 'First')
        for placement in ('document', 'inline', None, 0):
            do(METHODS[1], {'code': 'NOTICE', 'message': 'one', 'placement': placement})
    elif case == 'unhashable-placement-before-message':
        do(METHODS[0])
        do(METHODS[1], AuditedDict(events, code='NOTICE', placement=[], message=Poison(events, 'unreachable message')))
    elif case == 'missing-anchor-before-issue':
        do(METHODS[1], AuditedIssue(events, code='NOTICE'))
    elif case == 'reserve-failure-before-notices':
        native.fail_reservation = True
        do(METHODS[0], 'First', [AuditedIssue(events, code='NOTICE')])
    elif case == 'title-coercion-before-reserve':
        do(METHODS[0], Poison(events, 'title'), [])
    elif case == 'serial-native-failure':
        native.fail_box = 2
        do(METHODS[2], [{'code': 'ONE', 'message': 'first'}, {'code': 'TWO', 'message': 'second'}, {'code': 'THREE', 'message': 'never'}])
    elif case == 'serial-iterator-failure':
        def notices():
            yield {'code': 'ONE', 'message': 'first'}
            events.append(('advance', 'second'))
            raise RuntimeError('iterator failed')
        do(METHODS[0], 'First', notices())
    elif case == 'failed-first-notice-retry-title':
        do(METHODS[0], 'First')
        native.fail_box = 1
        do(METHODS[1], {'code': 'NOTICE', 'message': 'one'})
        native.fail_box = None
        do(METHODS[1], {'code': 'NOTICE', 'message': 'one'})
    elif case == 'private-redaction-and-unicode':
        do(METHODS[0], '/Users/customer/private/title.txt')
        for node in ('/Users/customer/private/a.txt', '/Users/customer/private/b.txt'):
            do(METHODS[1], {'code': 'NOTICE', 'nodeId': node, 'message': '中文😀 /Users/customer/private/secret.txt'})
    elif case == 'saved-cursor-not-public-bookmark':
        do(METHODS[0], 'First')
        do(METHODS[1], {'code': 'NOTICE', 'message': 'one'})
        do(METHODS[3], code='PATCH', message='near default', fallback='notice-only', node_id=Poison(events, 'unused node id'), page=1)
    elif case in ('explicit-first-paragraph', 'numeric-bookmark', 'falsy-default-bookmark', 'bookmark-coercion'):
        do(METHODS[0], 'First')
        name = 2 if case == 'numeric-bookmark' else '' if case == 'falsy-default-bookmark' else 'explicit'
        do(METHODS[3], code='bad code', message=None, fallback=None, node_id=Poison(events, 'unused node id'), page=page, bookmark_name=name)
        if case == 'explicit-first-paragraph':
            do(METHODS[3], code='bad code', message=None, fallback=None, node_id=Poison(events, 'unused node id'), page=page, bookmark_name=name)
    elif case == 'missing-bookmark-before-page':
        do(METHODS[3], code='PATCH', message=Poison(events, 'message'), fallback='fallback', node_id=None, page=Poison(events, 'page'), bookmark_name='missing')
    elif case == 'bookmark-message-before-page':
        do(METHODS[3], code='PATCH', message=Poison(events, 'message'), fallback=Poison(events, 'fallback'), node_id=None, page=Poison(events, 'page'), bookmark_name='explicit')
    elif case == 'target-error-before-display':
        do(METHODS[0])
        native.range_failure = RuntimeError('native range failed')
        do(METHODS[1], {'code': 'NOTICE', 'message': Poison(events, 'display')})
    else:
        raise AssertionError('Unknown contract case '+case)
    return capture(obj, native, results)


CASES = ('repeated-empty-reserve', 'first-title-cursor-and-dedup', 'dict-object-distinction',
         'empty-fallback-uses-message', 'placement-is-identity-not-layout', 'unhashable-placement-before-message',
         'missing-anchor-before-issue', 'reserve-failure-before-notices', 'title-coercion-before-reserve',
         'serial-native-failure', 'serial-iterator-failure', 'failed-first-notice-retry-title',
         'private-redaction-and-unicode', 'saved-cursor-not-public-bookmark', 'explicit-first-paragraph',
         'numeric-bookmark', 'falsy-default-bookmark', 'missing-bookmark-before-page',
         'bookmark-message-before-page', 'target-error-before-display')


@pytest.mark.parametrize('case', CASES)
def test_quality_semantics_match_frozen_writer(case, monkeypatch):
    expected = scenario('writer', case, monkeypatch)
    actual = scenario('mac', case, monkeypatch)
    assert actual == expected


@pytest.mark.parametrize('page', [0, -2, '7', 2.9, True, False, None, 'bad'])
def test_bookmark_page_coercion_and_failure_order_match_frozen_writer(page, monkeypatch):
    assert scenario('mac', 'bookmark-coercion', monkeypatch, page) == scenario('writer', 'bookmark-coercion', monkeypatch, page)


@pytest.mark.parametrize('name', METHODS)
def test_public_signatures_keep_frozen_positional_and_keyword_only_contract(name):
    actual = getattr(MacWordSession, name, None)
    assert callable(actual), 'Direct Mac Word quality method is missing: '+name
    assert inspect.signature(actual) == inspect.signature(getattr(WriterComposer, name))


def test_frozen_oracle_first_title_cursor_and_duplicate_is_meaningful(monkeypatch):
    result = scenario('writer', 'first-title-cursor-and-dedup', monkeypatch)
    assert result['results'] == [('return', None)] * 4
    assert [e for e in result['events'] if e[0] == 'box'] == [('box', 9, 'First\r[NOTICE] one'), ('box', 29, '[NOTICE] two')]
    assert result['position'] == 43 and result['bookmark'] == (9, 9)
    assert result['seen'] == {('NOTICE', 'document', 'n1'), ('NOTICE', 'document', 'n2')}
    assert [e for e in result['events'] if e[0] == 'get'] == [('get', 'code'), ('get', 'placement'), ('get', 'nodeId')]


def test_frozen_oracle_serial_failure_preserves_first_commit(monkeypatch):
    result = scenario('writer', 'serial-native-failure', monkeypatch)
    assert result['results'] == [('error', 'NativeWriterObjectError', 'DEGRADATION_INSERT_FAILED', 'quality notice upsert failed')]
    assert result['seen'] == {('ONE', 'document', '')}
    assert [e[2] for e in result['events'] if e[0] == 'box'] == ['生成质量提示\r[ONE] first', '[TWO] second']
    assert result['position'] == 29


def test_frozen_oracle_bookmark_uses_first_paragraph_end_and_never_deduplicates(monkeypatch):
    result = scenario('writer', 'explicit-first-paragraph', monkeypatch)
    assert [e for e in result['events'] if e[0] == 'box'] == [('box', 31, '[DEGRADATION] None (page 1; None)')] * 2
    assert result['position'] == 9 and result['seen'] == set()
    assert not [e for e in result['events'] if e[0] == 'consume']


@pytest.mark.parametrize(('case', 'error', 'events'), [
    ('title-coercion-before-reserve', ('error', 'ValueError', None, 'poison title'), [('consume', 'title')]),
    ('missing-anchor-before-issue', ('error', 'NativeWriterObjectError', 'DEGRADATION_INSERT_FAILED', 'quality anchor is unavailable'), []),
    ('missing-bookmark-before-page', ('error', 'NativeWriterObjectError', 'DEGRADATION_INSERT_FAILED', 'quality notice insertion failed'), [('bookmark', 'missing')]),
])
def test_frozen_oracle_errors_happen_before_later_consumption(case, error, events, monkeypatch):
    result = scenario('writer', case, monkeypatch)
    assert result['results'] == [error]
    assert result['events'] == events
