"""Frozen Writer differential contracts; native ports are doubles, not native proof."""
from __future__ import annotations

import ast
from copy import deepcopy
import importlib
import inspect
import re
import subprocess
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_fields import NativeIndexHandle
from skills.WPSComposer.scripts.msoffice.errors import NativeWordError, NativeWordTimeoutError

BOOKMARK = 'wpsc_eq_' + 'a' * 24
GLOBAL = dict(sequenceId='WPSC_EQ', mode='global', prefix='(', suffix=')')
CHAPTER = dict(GLOBAL, mode='chapter', chapterStyleLevel=1, resetLevel=1)


def implementation():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_numbering')
    except ModuleNotFoundError:
        pytest.fail('Bound native numbering implementation is missing')


@pytest.fixture(scope='module')
def frozen_writer():
    source = subprocess.check_output(['git', 'show', '6dd3a00:skills/WPSComposer/scripts/writer.py'], text=True)
    tree = ast.parse(source)
    methods = {'_native_fields', '_native_position', '_native_insert_field',
               '_localized_styleref_code', '_add_native_number_shell', '_add_native_caption',
               'add_equation_number_native', 'snapshot_fields'}
    cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'WriterComposer')
    cls.bases = []
    cls.body = [n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name in methods]
    nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_caption_field_codes']
    namespace = dict(__name__='skills.WPSComposer.scripts._frozen_writer', __package__='skills.WPSComposer.scripts', re=re, _NATIVE_SEQUENCE_IDS={'WPSC_FIG', 'WPSC_TAB', 'WPSC_EQ'},
                     _NATIVE_BOOKMARK_RE=re.compile(r'^wpsc_(?:fig|tab|eq)_[0-9a-f]{24}$'))
    exec(compile(ast.Module(body=nodes+[cls], type_ignores=[]), '<frozen-6dd3a00>', 'exec'), namespace)
    return namespace['WriterComposer']


class NativeModel:
    """COM/AppleEvent boundary model. Logical orchestration remains real code."""
    def __init__(self, start=4, end=4, style='标题 1', failure=None):
        self.start, self.position, self.style, self.failure = start, end, style, failure
        self.events = []
        self.fields = []

    def text(self, text):
        if not isinstance(text, str):
            raise TypeError('native text must be text')
        self.events.append(('text', self.start, self.position, text))
        self.position = self.start + len(text.encode('utf-16-le')) // 2
        self.start = self.position

    def field(self, code, owner, kind):
        self.events.append(('field', self.position, code))
        if self.failure == 'field':
            raise RuntimeError('native field failed')
        self.position += 3  # Simulated result + field end, no rendered-number assumption.
        self.start = self.position
        self.fields.append((owner or 'doc:native', kind, 'numbering'))

    def localized_style(self):
        self.events.append(('style',))
        if isinstance(self.style, Exception):
            return 'Heading 1'
        return self.style

    def bookmark(self, name, start, end):
        self.events.append(('bookmark', name, start, end))
        if self.failure == 'bookmark':
            raise RuntimeError('native bookmark failed')

    def format_paragraph(self, start):
        self.events.append(('format', start, self.position, 2, -1))
        if self.failure == 'format':
            raise RuntimeError('native paragraph failed')

    def format_caption(self, start, keep_with_next):
        self.events.append(('format', start, self.position, 1, -1))
        if self.failure == 'format':
            raise RuntimeError('native paragraph failed')
        value = -1 if keep_with_next else 0
        self.events.append(('keep-next', value))
        if self.failure == 'keep':
            raise RuntimeError('native keep failed')
        return SimpleNamespace(Start=start, End=self.position)

    def paragraph(self):
        self.events.append(('paragraph', self.position))
        self.position += 1
        self.start = self.position


def oracle(cls, model):
    writer = cls()
    class Selection:
        @property
        def Range(self):
            return SimpleNamespace(Start=model.start, End=model.position, Collapse=lambda direction: None)
        def TypeText(self, text): model.text(text)
        def TypeParagraph(self): model.paragraph()
        def SetRange(self, start, end): model.start, model.position = start, end
        def MoveRight(self, unit, count): model.start += 1; model.position += 1
    class Fields:
        def Add(self, rng, field_type, code, preserve):
            # Tracking happens in real frozen _native_insert_field, not here.
            model.events.append(('field', rng.End, code))
            if model.failure == 'field': raise RuntimeError('native field failed')
            return SimpleNamespace(Result=SimpleNamespace(End=rng.End+2))
    class Format:
        def __init__(self, start, end): self.start, self.end = start, end
        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == 'KeepTogether':
                model.events.append(('format', self.start, self.end, self.Alignment, value))
                if model.failure == 'format': raise RuntimeError('native paragraph failed')
            elif name == 'KeepWithNext':
                model.events.append(('keep-next', value))
    def style(index):
        model.events.append(('style',))
        if isinstance(model.style, Exception): raise model.style
        return SimpleNamespace(NameLocal=model.style)
    selection = Selection()
    try:
        writer.selection = selection
    except AttributeError:
        writer._app = SimpleNamespace(Selection=selection)
    writer._doc = SimpleNamespace(Fields=Fields(), Styles=style,
        Range=lambda start,end: SimpleNamespace(Start=start, End=end, ParagraphFormat=Format(start,end)),
        Bookmarks=SimpleNamespace(Add=lambda name,rng: model.bookmark(name,rng.Start,rng.End)))
    return writer


def outcome(call):
    try: return ('return', call())
    except Exception as error: return ('error', type(error), str(error))


@pytest.mark.parametrize('numbering,source,fallback,bookmark,style,start,end,failure', [
    (GLOBAL, 'x=1', 'unused', BOOKMARK, '标题 1', 4, 4, None),
    (CHAPTER, '中文😀', 'unused', BOOKMARK, '标题 1', 4, 4, None),
    (CHAPTER, 0, 7, None, 'Heading 1', 4, 9, None),
    (CHAPTER, False, None, BOOKMARK, '', 4, 4, None),
    (CHAPTER, '', 'fallback', BOOKMARK, 'bad"name', 4, 4, None),
    (CHAPTER, '', 'fallback', BOOKMARK, RuntimeError('style unavailable'), 4, 4, None),
    (dict(GLOBAL,prefix='😀(',suffix=')尾'), [], False, BOOKMARK, '标题 1', 4, 4, None),
    (dict(GLOBAL,sequenceId='WPSC_TAB'), 'x', '', BOOKMARK, '标题 1', 0, 0, None),
    (dict(CHAPTER,sequenceId='WPSC_FIG'), 'x', '', BOOKMARK, '标题 1', 0, 0, None),
    ({}, 'already inserted', '', 'bad', '标题 1', 4, 4, None),
    (dict(GLOBAL,resetLevel=1), 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    (dict(CHAPTER,chapterStyleLevel=2), 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    ({k:v for k,v in GLOBAL.items() if k!='suffix'}, 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', 'bad', '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', '', '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', 42, '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'field'),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'bookmark'),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'format'),
])
def test_equation_preserves_frozen_order_and_partial_failure(monkeypatch,frozen_writer,
        numbering,source,fallback,bookmark,style,start,end,failure):
    expected = NativeModel(start,end,style,failure)
    baseline = oracle(frozen_writer,expected)
    kwargs = dict(source=source,numbering=numbering,bookmarkName=bookmark,
                  fallbackText=fallback,owner_node_id='node:equation')
    want = outcome(lambda:baseline.add_equation_number_native(**kwargs))
    actual = NativeModel(start,end,style,failure)
    m = implementation()
    monkeypatch.setattr(m,'BoundNumberingCursor',lambda session:actual)
    got = outcome(lambda:MacWordSession().add_equation_number_native(**kwargs))
    assert got == want
    assert actual.events == expected.events
    assert actual.fields == [(owner,kind,category) for owner,kind,field,category in baseline._native_fields()]


def test_public_signature_matches_frozen(frozen_writer):
    assert hasattr(MacWordSession,'add_equation_number_native')
    assert inspect.signature(MacWordSession.add_equation_number_native) == inspect.signature(frozen_writer.add_equation_number_native)


@pytest.mark.parametrize('numbering,caption,bookmark,style,start,end,keep_with_next', [
    (dict(GLOBAL, sequenceId='WPSC_FIG', prefix='图😀(', suffix=')尾'),
     '说明😀', BOOKMARK.replace('eq', 'fig'), '标题 1', 4, 4, False),
    (dict(CHAPTER, sequenceId='WPSC_TAB', prefix='😀[', suffix=']尾'),
     True, BOOKMARK.replace('eq', 'tab'), '标题 1', 4, 4, True),
    (dict(GLOBAL, sequenceId='WPSC_FIG'), 0, None, 'Heading 1', 7, 7, False),
    (dict(GLOBAL, sequenceId='WPSC_FIG'), [], None, 'Heading 1', 7, 7, True),
])
def test_caption_preserves_frozen_order_truthiness_affixes_range_and_keep_flags(
        monkeypatch, frozen_writer, numbering, caption, bookmark, style, start, end,
        keep_with_next):
    expected = NativeModel(start, end, style)
    baseline = oracle(frozen_writer, expected)
    args = (caption, numbering, bookmark, 'node:caption')
    want = outcome(lambda: baseline._add_native_caption(
        *args, keep_with_next=keep_with_next))

    actual = NativeModel(start, end, style)
    m = implementation()
    monkeypatch.setattr(m, 'BoundNumberingCursor', lambda session: actual)
    got = outcome(lambda: MacWordSession()._add_native_caption(
        *args, keep_with_next=keep_with_next))

    assert got[0] == want[0] == 'return'
    assert (got[1].Start, got[1].End) == (want[1].Start, want[1].End)
    assert actual.events == expected.events
    assert actual.fields == [
        (owner, kind, category)
        for owner, kind, field, category in baseline._native_fields()
    ]


def test_private_caption_signature_matches_frozen(frozen_writer):
    assert inspect.signature(MacWordSession._add_native_caption) == inspect.signature(
        frozen_writer._add_native_caption)


def test_caption_evaluates_keep_truthiness_after_base_format_like_frozen(
        monkeypatch, frozen_writer):
    class LateTruthError:
        def __bool__(self):
            raise ValueError('late keep truthiness')

    expected = NativeModel()
    baseline = oracle(frozen_writer, expected)
    keep = LateTruthError()
    args = ('caption', dict(GLOBAL, sequenceId='WPSC_FIG'), None, 'fig:order')
    want = outcome(lambda: baseline._add_native_caption(
        *args, keep_with_next=keep))

    actual = NativeModel()
    module = implementation()
    monkeypatch.setattr(module, 'BoundNumberingCursor', lambda session: actual)
    got = outcome(lambda: MacWordSession()._add_native_caption(
        *args, keep_with_next=keep))

    assert got == want == ('error', ValueError, 'late keep truthiness')
    assert actual.events == expected.events


def test_caption_base_format_failure_precedes_keep_truthiness_like_frozen(
        monkeypatch, frozen_writer):
    class MustNotEvaluate:
        def __bool__(self):
            raise AssertionError('keep evaluated too early')

    expected = NativeModel(failure='format')
    baseline = oracle(frozen_writer, expected)
    args = ('caption', dict(GLOBAL, sequenceId='WPSC_FIG'), None, 'fig:order')
    want = outcome(lambda: baseline._add_native_caption(
        *args, keep_with_next=MustNotEvaluate()))

    actual = NativeModel(failure='format')
    module = implementation()
    monkeypatch.setattr(module, 'BoundNumberingCursor', lambda session: actual)
    got = outcome(lambda: MacWordSession()._add_native_caption(
        *args, keep_with_next=MustNotEvaluate()))

    assert got == want == ('error', RuntimeError, 'native paragraph failed')
    assert actual.events == expected.events


def test_windows_writer_caption_uses_original_start_for_long_replacement():
    from skills.WPSComposer.scripts.writer import WriterComposer

    model = NativeModel(start=4, end=100)
    writer = oracle(WriterComposer, model)
    result = writer._add_native_caption(
        '说明😀', dict(GLOBAL, sequenceId='WPSC_FIG', prefix='图😀(', suffix=')尾'),
        BOOKMARK.replace('eq', 'fig'), 'fig:windows')

    assert (result.Start, result.End) == (4, 18)
    assert ('text', 11, 11, ')尾') in model.events


def test_numbering_fields_have_owner_category_and_per_kind_ordinal(monkeypatch):
    s = MacWordSession()
    codes = [' STYLEREF "标题 1" \\s ', ' SEQ WPSC_EQ \\* ARABIC \\s 1 ', ' SEQ WPSC_EQ \\* ARABIC ']
    kinds = ['STYLEREF','SEQ_EQ','SEQ_EQ']
    s._tracked_numbering = [(NativeIndexHandle('s',f'own{i}','eq:one',kind,'numbering'),code)
                           for i,(kind,code) in enumerate(zip(kinds,codes))]
    rows = [['stats',1]]
    for i,(kind,code) in enumerate(zip(kinds,codes)):
        rows += [['identity',f'own{i}',i*50+1], ['field','story:main text/chain:1',i+1,kind,code,i*50+1,str(i+1),0,0,1]]
    monkeypatch.setattr(s,'_execute',lambda lines:deepcopy(rows))
    got = s.snapshot_fields()
    assert [f.stable_key for f in got] == [('eq:one','STYLEREF',0),('eq:one','SEQ_EQ',0),('eq:one','SEQ_EQ',1)]
    assert all(f.field_category=='numbering' for f in got)
    rows[2][4] = ' SEQ WPSC_TAB \\* ARABIC '
    with pytest.raises(NativeWordError): s.snapshot_fields()


def test_checkpoint_restores_numbering_tracking_after_confirmed_rollback(monkeypatch):
    from tests.msoffice.test_macos_word_recovery import rows,bind
    a = rows()
    field = ['field','main',1,'SEQ_EQ',' SEQ WPSC_EQ \\* ARABIC ',8,33,34,35]
    s,t = bind(monkeypatch,[a,rows(end=40,fields=[field]),[['rollback-ack',[field],[]],*a]])
    token = s.degradation_checkpoint()
    s._tracked_numbering = [(NativeIndexHandle('s','new','n','SEQ_EQ','numbering'),field[4])]
    assert s.rollback_degradation_checkpoint(token) is None
    assert s._tracked_numbering == []


def test_numbering_tracking_prefix_cannot_be_rewritten_before_rollback(monkeypatch):
    from tests.msoffice.test_macos_word_recovery import rows,bind
    handle = NativeIndexHandle('s','old','n','SEQ_EQ','numbering')
    a = rows(bookmarks=[['bookmark','old',1,4,'abc']])
    s,t = bind(monkeypatch,[a])
    s._tracked_numbering = [(handle,'abc')]
    token = s.degradation_checkpoint()
    s._tracked_numbering[0] = (handle,'rewritten')
    with pytest.raises(Exception) as error: s.rollback_degradation_checkpoint(token)
    assert error.value.code == 'LOCAL_MUTATION_ROLLBACK_FAILED'
    assert len(t.calls) == 1


def bind_cursor(monkeypatch, replies):
    m = implementation()
    s = MacWordSession()
    s._bound_path = '/private/owned.docx'
    calls = []
    def execute(lines):
        calls.append(lines)
        row = replies.pop(0)
        if isinstance(row,BaseException):
            # This double represents submitted native errors. Local script-I/O
            # failures are tested separately through the real transport below.
            s._invalidate_field_topology()
            s._field_topology_mutation_pending=False
            raise row
        return deepcopy(row)
    monkeypatch.setattr(s,'_execute',execute)
    monkeypatch.setattr(m,'uuid4',lambda:SimpleNamespace(hex='a'*32))
    return s,calls,m


def field_row():
    return ['field',2,3,10,35,36,37,' SEQ WPSC_EQ \\* ARABIC ',
            'WPSC_N_'+'a'*30,10,35,' SEQ WPSC_EQ \\* ARABIC ']


def field_row_at(position, code, *, before=2):
    identity = 'WPSC_N_' + 'a' * 30
    code_start = position + 1
    code_end = code_start + 2
    result_start = code_end + 1
    result_end = result_start + 1
    return [
        'field', before, before + 1, code_start, code_end,
        result_start, result_end, f' {code} ', identity,
        code_start, code_end, f' {code} ',
    ]


def test_real_caption_cursor_replaces_long_middle_selection_and_acks_exact_range(monkeypatch):
    numbering = dict(
        GLOBAL, sequenceId='WPSC_FIG', prefix='图😀(', suffix=')尾')
    code = 'SEQ WPSC_FIG \\* ARABIC'
    replies = [
        [['selection', 4, 100]],
        [['text', 4, 8, '图😀(']],
        [field_row_at(8, code)],
        [['bookmark', BOOKMARK.replace('eq', 'fig'), 8, 14]],
        [['text', 14, 16, ')尾']],
        [['text', 16, 21, ' 说明😀']],
        [['caption-format-base', 4, 21, True, True]],
        [['caption-format-keep', 4, 21, False]],
        [['text', 21, 22, '\r']],
    ]
    session, calls, module = bind_cursor(monkeypatch, replies)

    result = session._add_native_caption(
        '说明😀', numbering, BOOKMARK.replace('eq', 'fig'), 'fig:one')

    assert (result.Start, result.End) == (4, 21)
    assert not replies
    handle, tracked_code = session._tracked_numbering[0]
    assert (handle.owner_node_id, handle.kind, handle.category) == (
        'fig:one', 'SEQ_FIG', 'numbering')
    assert tracked_code == f' {code} '
    base_format_call, keep_call = calls[-3:-1]
    assert 'set alignment of paragraph format of numberRange to align paragraph center' in base_format_call
    assert 'set keep together of paragraph format of numberRange to true' in base_format_call
    assert 'set keep with next of paragraph format of numberRange to false' in keep_call
    assert not any(term in '\n'.join(base_format_call + keep_call) for term in (
        'first line indent', 'left indent', 'right indent', 'space before',
        'space after', 'line spacing', 'font object'))


def test_real_chapter_caption_uses_styleref_unicode_affixes_and_true_keep(monkeypatch):
    numbering = dict(
        CHAPTER, sequenceId='WPSC_TAB', prefix='表😀[', suffix=']尾')
    style_code = 'STYLEREF "标题 1" \\s'
    sequence_code = 'SEQ WPSC_TAB \\* ARABIC \\s 1'
    replies = [
        [['selection', 12, 12]],
        [['text', 12, 16, '表😀[']],
        [['style', '标题 1']],
        [field_row_at(16, style_code)],
        [['text', 22, 23, '-']],
        [field_row_at(23, sequence_code, before=3)],
        [['bookmark', BOOKMARK.replace('eq', 'tab'), 16, 29]],
        [['text', 29, 31, ']尾']],
        [['caption-format-base', 12, 31, True, True]],
        [['caption-format-keep', 12, 31, True]],
        [['text', 31, 32, '\r']],
    ]
    session, calls, module = bind_cursor(monkeypatch, replies)

    result = session._add_native_caption(
        0, numbering, BOOKMARK.replace('eq', 'tab'), 'tab:one',
        keep_with_next=True,
    )

    assert (result.Start, result.End) == (12, 31)
    assert not replies
    assert [(h.kind, code) for h, code in session._tracked_numbering] == [
        ('STYLEREF', f' {style_code} '),
        ('SEQ_TAB', f' {sequence_code} '),
    ]
    assert 'set keep with next of paragraph format of numberRange to true' in calls[-2]


def test_caption_late_bookmark_error_retains_confirmed_field_without_suffix(monkeypatch):
    numbering = dict(GLOBAL, sequenceId='WPSC_FIG')
    code = 'SEQ WPSC_FIG \\* ARABIC'
    session, calls, module = bind_cursor(monkeypatch, [
        [['selection', 4, 4]],
        [['text', 4, 5, '(']],
        [field_row_at(5, code)],
    ])

    with pytest.raises(ValueError, match='invalid native bookmark'):
        session._add_native_caption('never written', numbering, 'bad', 'fig:late')

    assert len(session._tracked_numbering) == 1
    assert len(calls) == 3
    assert not session._quarantined


@pytest.mark.parametrize('bad', [
    [['caption-format-base', 0, True, True, True]],
    [['caption-format-base', 0, 1, True, False]],
])
def test_caption_format_requires_strict_typed_ack_and_quarantines(monkeypatch, bad):
    session, calls, module = bind_cursor(monkeypatch, [
        [['selection', 0, 1]], bad,
    ])
    cursor = module.BoundNumberingCursor(session)

    with pytest.raises(NativeWordError):
        cursor.format_caption(0, False)

    assert session._quarantined and session._retain_evidence
    assert len(calls) == 2


@pytest.mark.parametrize('bad', [
    [['caption-format-keep', 0, 1, 0]],
    [['caption-format-keep', 0, True, False]],
])
def test_caption_keep_requires_strict_typed_ack_and_quarantines(monkeypatch, bad):
    session, calls, module = bind_cursor(monkeypatch, [
        [['selection', 0, 1]],
        [['caption-format-base', 0, 1, True, True]],
        bad,
    ])
    cursor = module.BoundNumberingCursor(session)

    with pytest.raises(NativeWordError):
        cursor.format_caption(0, False)

    assert session._quarantined and session._retain_evidence
    assert len(calls) == 3


def test_caption_keep_truthiness_is_evaluated_after_base_format_ack(monkeypatch):
    class LateTruthError:
        def __bool__(self):
            raise ValueError('late keep truthiness')

    session, calls, module = bind_cursor(monkeypatch, [
        [['selection', 4, 100]],
        [['caption-format-base', 4, 21, True, True]],
    ])
    cursor = module.BoundNumberingCursor(session)
    cursor.start = cursor.position = 21

    with pytest.raises(ValueError, match='late keep truthiness'):
        cursor.format_caption(4, LateTruthError())

    assert len(calls) == 2
    assert not session._quarantined


def test_caption_rejects_reversed_range_before_format_submission(monkeypatch):
    session, calls, module = bind_cursor(monkeypatch, [[['selection', 4, 100]]])
    cursor = module.BoundNumberingCursor(session)
    cursor.start = cursor.position = 21

    with pytest.raises(NativeWordError):
        cursor.format_caption(100, False)

    assert len(calls) == 1
    assert session._quarantined


def test_caption_submitted_operation_failure_quarantines_without_retry(monkeypatch):
    session, calls, module = bind_cursor(monkeypatch, [
        [['selection', 4, 4]], RuntimeError('submitted caption failure'),
    ])

    with pytest.raises(RuntimeError, match='submitted caption failure'):
        session._add_native_caption(
            'caption', dict(GLOBAL, sequenceId='WPSC_FIG'), None, 'fig:failure')

    assert session._quarantined and session._retain_evidence
    assert len(calls) == 2


def test_real_cursor_consumes_exact_noncollapsed_selection_and_native_field_end(monkeypatch):
    replies = [[['selection',4,9]],[['text',4,7,'x😀']],[['text',7,8,'\t']],
               [['text',8,9,'(']],[field_row()],
               [['bookmark',BOOKMARK,9,38]],[['text',38,39,')']],
               [['format',True,True]],[['text',39,40,'\r']]]
    s,calls,m = bind_cursor(monkeypatch,replies)
    s._observed_field_topology = (('before',),)
    assert s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=BOOKMARK,
                                        fallbackText='unused') == {'issues':[]}
    assert not replies
    assert s._structural_changed and not hasattr(s,'_observed_field_topology')
    handle,code = s._tracked_numbering[0]
    assert (handle.owner_node_id,handle.kind,handle.category)==('doc:native','SEQ_EQ','numbering')
    assert code == ' SEQ WPSC_EQ \\* ARABIC '


@pytest.mark.parametrize('change', ['count','position','bounds','code','identity','identity-bounds','identity-text'])
def test_malformed_native_field_ack_never_commits_tracking(monkeypatch,change):
    row = field_row()
    if change=='count': row[2]=4
    elif change=='position': row[3]=11
    elif change=='bounds': row[5]=34
    elif change=='code': row[7]=row[11]=' SEQ ATTACK \\* ARABIC '
    elif change=='identity': row[8]='foreign'
    elif change=='identity-bounds': row[9]=11
    else: row[11]='different'
    s,calls,m = bind_cursor(monkeypatch,[[['selection',9,9]],[row]])
    c = m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.field('SEQ WPSC_EQ \\* ARABIC','owner','SEQ_EQ')
    assert s._quarantined and s._retain_evidence and not hasattr(s,'_tracked_numbering')


@pytest.mark.parametrize('bad', [[],[['text',4,7,'wrong']],[['text',4,8,'x😀']]])
def test_literal_bad_ack_quarantines_without_continuing(monkeypatch,bad):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',4,9]],bad])
    with pytest.raises(NativeWordError):
        s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=None,fallbackText='')
    assert s._quarantined and len(calls)==2


@pytest.mark.parametrize('error',[NativeWordTimeoutError(),NativeWordError('NATIVE_WORD_EXECUTION_FAILED'),KeyboardInterrupt()])
def test_uncertain_mutation_stops_without_retry(monkeypatch,error):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',4,4]],error])
    with pytest.raises(type(error)):
        s.add_equation_number_native(source='x',numbering=GLOBAL,bookmarkName=None,fallbackText='')
    assert s._quarantined and s._retain_evidence and len(calls)==2


@pytest.mark.parametrize('state',['_closed','_read_only','_quarantined'])
def test_unavailable_session_cannot_submit_any_native_call(monkeypatch,state):
    s,calls,m = bind_cursor(monkeypatch,[])
    setattr(s,state,True)
    with pytest.raises((ValueError,NativeWordError)):
        s.add_equation_number_native(source='x',numbering={},bookmarkName=None,fallbackText='')
    assert not calls


def test_late_bookmark_validation_preserves_confirmed_native_field_tracking(monkeypatch):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',6,6]],[['text',6,7,'x']],
        [['text',7,8,'\t']],[['text',8,9,'(']],[field_row()]])
    with pytest.raises(ValueError,match='invalid native bookmark'):
        s.add_equation_number_native(source='x',numbering=GLOBAL,bookmarkName='bad',fallbackText='')
    assert len(s._tracked_numbering)==1 and not s._quarantined
    assert len(calls)==5


def test_numbering_native_scripts_compile_without_execution(monkeypatch,tmp_path):
    import sys
    if sys.platform!='darwin': pytest.skip('Word dictionary compile only')
    replies = [[['selection',4,9]],[['text',4,7,'x😀']],[['text',7,8,'\t']],
               [['text',8,9,'(']],[field_row()],
               [['bookmark',BOOKMARK,9,38]],[['text',38,39,')']],
               [['format',True,True]],[['text',39,40,'\r']]]
    s,calls,m = bind_cursor(monkeypatch,replies)
    s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=BOOKMARK,fallbackText='')
    c = m.BoundNumberingCursor.__new__(m.BoundNumberingCursor)
    c.session=s; c.start=c.position=40
    replies.append([['style','标题 1']]); c.localized_style()
    replies.append([['caption-format-base',40,40,True,True]])
    replies.append([['caption-format-keep',40,40,False]])
    c.format_caption(40,False)
    for index,lines in enumerate(calls):
        path = tmp_path/f'{index}.applescript'
        path.write_text('use framework "Foundation"\nuse scripting additions\ntell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell\n')
        result = subprocess.run(['/usr/bin/osacompile','-o',str(path.with_suffix('.scpt')),str(path)],capture_output=True,text=True,timeout=20)
        assert result.returncode==0,result.stderr


def test_native_fixture_requires_explicit_execute_before_creating_directory(tmp_path):
    import sys
    result=subprocess.run([sys.executable,'fixtures/microsoft_parity/macos_word_numbering.py',
        '--output',str(tmp_path/'untouched')],capture_output=True,text=True)
    assert result.returncode==2 and '--execute' in result.stderr
    assert not (tmp_path/'untouched').exists()


def test_native_fixture_validates_number_only_bookmarks_and_real_fields():
    try: m=importlib.import_module('fixtures.microsoft_parity.macos_word_numbering')
    except ModuleNotFoundError: pytest.fail('Guarded native numbering fixture is missing')
    rows=[['body','Chapter\rleft x=1\t😀(1)尾\rfallback\t[1-1]\r7\t(2)\r right\rTAIL\r'],
          ['field','SEQ WPSC_EQ \\* ARABIC','1'],['field','STYLEREF "标题 1" \\s','1'],
          ['field','SEQ WPSC_FIG \\* ARABIC \\s 1','1'],['field','SEQ WPSC_EQ \\* ARABIC','2'],
          ['bookmark',m.BOOKMARKS[0],20,50,'1'],['bookmark',m.BOOKMARKS[1],60,100,'1-1'],
          ['bookmark',m.BOOKMARKS[2],110,140,'2'],
          ['counts',0,0,0],['paragraph','left x=1\t😀(1)尾\r',True,True]]
    assert m.validate_readback(rows)
    broken=deepcopy(rows);broken[5][-1]='😀(1)尾'
    assert not m.validate_readback(broken)
    broken=deepcopy(rows);broken[2][1]='REF fake'
    assert not m.validate_readback(broken)


def test_caption_fixture_requires_explicit_execute_before_creating_directory(tmp_path):
    import sys
    result=subprocess.run([sys.executable,'fixtures/microsoft_parity/macos_word_caption.py',
        '--output',str(tmp_path/'untouched')],capture_output=True,text=True)
    assert result.returncode==2 and '--execute' in result.stderr
    assert not (tmp_path/'untouched').exists()


def test_caption_fixture_rejects_malformed_sentinel_ack_and_quarantines(monkeypatch,tmp_path):
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    class Session:
        staging_root=None
        def __enter__(self): return self
        def __exit__(self,*args): return None
        def _execute(self,lines):
            if any('captionSentinel to make new document' in line for line in lines):
                return [[True]]
            return [['seed','1']]
        def _retain(self,reason): self.retained=reason
    session=Session()
    monkeypatch.setattr(module,'create_document',lambda *args,**kwargs:session)
    monkeypatch.setattr(module,'inventory',lambda *args:[])
    monkeypatch.setattr(module,'retain_sources',lambda *args:{})
    output=tmp_path/'malformed-sentinel'

    result=module.run(output,execute=True)

    assert result['status']=='FAIL'
    assert result['error']=={'type':'ValueError','message':'Invalid sentinel acknowledgement'}
    assert session.retained=='Native caption sentinel acknowledgement invalid'


def test_caption_fixture_validates_ranges_fields_centering_keep_and_suffix():
    try: m=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    except ModuleNotFoundError: pytest.fail('Guarded native caption fixture is missing')
    rows=[
        ['body','Chapter\rleft 图😀(1)尾 说明😀\r 表😀[1-1]尾 True\r\r\rTAIL\r\r'],
        ['field','SEQ WPSC_FIG \\* ARABIC','1'],
        ['field','STYLEREF "标题 1" \\s','1'],
        ['field','SEQ WPSC_TAB \\* ARABIC \\s 1','1'],
        ['bookmark',m.BOOKMARKS[0],17,60,'1'],
        ['bookmark',m.BOOKMARKS[1],73,161,'1-1'],
        ['paragraph',8,68,'left 图😀(1)尾 说明😀\r',True,True,False],
        ['paragraph',68,169,' 表😀[1-1]尾 True\r',True,True,True],
        ['counts',0,0,0],
    ]
    assert m.validate_readback(rows)
    for index,offset,replacement in (
        (4,-1,'图😀(1)尾'), (6,-3,False), (6,-2,False), (7,-1,False)):
        broken=deepcopy(rows)
        broken[index][offset]=replacement
        assert not m.validate_readback(broken)


def caption_native_rows(module):
    return [
        ['body','Chapter\rleft 图😀(1)尾 说明😀\r 表😀[1-1]尾 True\r\r\rTAIL\r\r'],
        ['field',' SEQ WPSC_FIG \\* ARABIC \\* MERGEFORMAT ','1'],
        ['field',' STYLEREF "标题 1" \\s \\* MERGEFORMAT ','1'],
        ['field',' SEQ WPSC_TAB \\* ARABIC \\s 1 \\* MERGEFORMAT ','1'],
        ['bookmark',module.BOOKMARKS[0],17,60,'1'],
        ['bookmark',module.BOOKMARKS[1],73,161,'1-1'],
        ['paragraph',8,68,'left 图😀(1)尾 说明😀\r',True,True,False],
        ['paragraph',68,169,' 表😀[1-1]尾 True\r',True,True,True],
        ['counts',0,0,0],
    ]


def test_caption_fixture_rejects_body_prefix_suffix_or_paragraph_corruption():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    rows=caption_native_rows(module)
    assert module.validate_readback(rows)
    for replacement in (
        rows[0][1].replace('Chapter\r','CORRUPTED\r',1),
        rows[0][1]+'AFTER TAIL',
        rows[0][1].replace('\r\r\rTAIL\r\r','\rTAIL\r'),
    ):
        broken=deepcopy(rows);broken[0][1]=replacement
        assert not module.validate_readback(broken)


def test_caption_fixture_rejects_foreign_or_missing_field_rows():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    rows=caption_native_rows(module)
    assert module.validate_readback(rows)
    foreign=deepcopy(rows);foreign.insert(4,['field',' DATE ','2026-09-11'])
    missing=deepcopy(rows);del missing[2]
    assert not module.validate_readback(foreign)
    assert not module.validate_readback(missing)


def test_caption_fixture_rejects_untyped_negative_or_inconsistent_bounds():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    rows=caption_native_rows(module)
    assert module.validate_readback(rows)
    for row_index,column,value in ((4,2,-10),(4,3,True),(6,1,-1),(6,2,9999)):
        broken=deepcopy(rows);broken[row_index][column]=value
        assert not module.validate_readback(broken)


def test_caption_fixture_validates_nondefault_format_and_trailing_paragraph():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    global_rows=[
        ['caption-format-detail','global',13,32,
         '图😀(1)尾 说明😀',True,True,False,17,3,5,11,7,21,True,14],
        ['caption-mark-detail','global',32,33,'\r'],
        ['following-text-detail','global',33,34,' '],
    ]
    chapter_rows=[
        ['caption-format-detail','chapter',34,52,
         '表😀[1-1]尾 True',True,True,True,17,3,5,11,7,21,True,14],
        ['caption-mark-detail','chapter',52,53,'\r'],
        ['following-text-detail','chapter',53,54,'\r'],
    ]
    assert module.validate_format_readback(
        global_rows,'global',start=13,end=32,following=' ')
    assert module.validate_format_readback(
        chapter_rows,'chapter',start=34,end=52,following='\r')
    for row_index,column,value in (
            (0,5,1), (0,8,0), (0,14,1), (1,2,True), (1,4,'changed'),
            (2,4,'right')):
        broken=deepcopy(global_rows);broken[row_index][column]=value
        assert not module.validate_format_readback(
            broken,'global',start=13,end=32,following=' ')


def test_caption_fixture_format_validator_binds_exact_handle_and_unit_bounds():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    rows=[
        ['caption-format-detail','global',13,32,
         '图😀(1)尾 说明😀',True,True,False,17,3,5,11,7,21,True,14],
        ['caption-mark-detail','global',32,33,'\r'],
        ['following-text-detail','global',33,34,' '],
    ]
    for row_index,bounds in ((0,[12,32]),(1,[32,38]),(2,[38,45])):
        broken=deepcopy(rows);broken[row_index][2:4]=bounds
        assert not module.validate_format_readback(
            broken,'global',start=13,end=32,following=' ')


def test_caption_fixture_pdf_validator_allows_only_local_emoji_extractor_order():
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    extracted='1 Chapter\nleft 😀 图 (1)尾 说 😀 明\n😀表 [1-1]尾 True\nTAIL'
    assert module.validate_pdf_text(extracted)
    for broken in (
            extracted.replace('😀','',1),
            extracted.replace('TAIL','TAIL😀'),
            extracted.replace('(1)尾','尾(1)'),
            extracted.replace('[1-1]尾','[1-1]'),
            extracted.replace('\nTAIL','\nEXTRA\nTAIL')):
        assert not module.validate_pdf_text(broken)


def test_caption_fixture_generated_applescripts_compile_without_execution(tmp_path):
    import sys
    if sys.platform!='darwin': pytest.skip('Word dictionary compile only')
    module=importlib.import_module('fixtures.microsoft_parity.macos_word_caption')
    groups=(
        module.seed_commands(),
        module.apply_nondefault_format_commands(module.LONG_START,module.LONG_END),
        module.format_readback_commands('global',13,32),
        module.readback_commands(),
    )
    for index,lines in enumerate(groups):
        path=tmp_path/f'caption-fixture-{index}.applescript'
        path.write_text(
            'use framework "Foundation"\nuse scripting additions\n'
            'tell application "Microsoft Word"\n'
            'set boundDoc to active document\nset boundWindow to active window\n'
            +'\n'.join(lines)+'\nend tell\n')
        result=subprocess.run(['/usr/bin/osacompile','-o',str(path.with_suffix('.scpt')),
                               str(path)],capture_output=True,text=True)
        assert result.returncode==0,result.stderr


from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_fields import NativeIndexHandle
from skills.WPSComposer.scripts.longform.field_contract import snapshot_visible_field


def test_main_numbering_does_not_collide_with_header(monkeypatch):
    s = MacWordSession()
    code = ' STYLEREF "Heading 1" \\s '
    s._tracked_numbering = [(NativeIndexHandle('s','own','eq:one','STYLEREF','numbering'),code)]
    monkeypatch.setattr(s, '_execute', lambda lines: [
        ['stats',1], ['identity','own',10],
        ['field','story:main text/chain:1',1,'STYLEREF',code,10,'1',0,0,1],
        ['field','story:primary header/chain:1',1,'STYLEREF',code,10,'1',0,0,1],
    ])
    result = s.snapshot_fields()
    assert [r.stable_key for r in result] == [
        ('eq:one','STYLEREF',0), ('story:primary header/chain:1','STYLEREF',0),
    ]


def test_numbering_snapshot_preserves_frozen_creation_order_ordinals(monkeypatch):
    s = MacWordSession()
    code = ' SEQ WPSC_EQ \\* ARABIC '
    s._tracked_numbering = [
        (NativeIndexHandle('s','first','eq:one','SEQ_EQ','numbering'),code),
        (NativeIndexHandle('s','second','eq:one','SEQ_EQ','numbering'),code),
    ]
    monkeypatch.setattr(s, '_execute', lambda lines: [
        ['stats',1], ['identity','first',100], ['identity','second',10],
        ['field','story:main text/chain:1',1,'SEQ_EQ',code,10,'SECOND',0,0,6],
        ['field','story:main text/chain:1',2,'SEQ_EQ',code,100,'FIRST',0,0,5],
    ])
    by_key = {r.stable_key:r.result_hash for r in s.snapshot_fields()}
    for ordinal,visible in enumerate(('FIRST','SECOND')):
        expected = snapshot_visible_field(owner_node_id='eq:one',field_kind='SEQ_EQ',
            ordinal_within_node=ordinal,visible_result=visible,field_category='numbering',
            toc_page_count=0,figure_index_page_count=0,table_index_page_count=0,total_pages=1)
        assert by_key[expected.stable_key] == expected.result_hash
    # Topology must retain native document order independently of semantic ordinals.
    assert [r[4] for r in s._observed_field_topology] == [10,100]


from skills.WPSComposer.scripts.msoffice.errors import NativeWordError
import pytest

@pytest.mark.parametrize("rows", [[None],[3],[True]])
def test_malformed_selection_ack_quarantines(monkeypatch,rows):
    s,calls,m=bind_cursor(monkeypatch,[rows])
    with pytest.raises(NativeWordError): m.BoundNumberingCursor(s)
    assert s._quarantined

def test_bool_coordinate_is_not_text_ack(monkeypatch):
    s,calls,m=bind_cursor(monkeypatch,[[["selection",0,0]],[["text",0,True,"x"]]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.text("x")
    assert s._quarantined

@pytest.mark.parametrize("kind", ["field","style"])
def test_malformed_postmutation_ack_quarantines(monkeypatch,kind):
    s,calls,m=bind_cursor(monkeypatch,[[["selection",0,0]],[None]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError):
        if kind=="field": c.field("SEQ WPSC_EQ \\* ARABIC","n","SEQ_EQ")
        else: c.localized_style()
    assert s._quarantined


@pytest.mark.parametrize('owner',['doc:native','eq:one'])
@pytest.mark.parametrize('kind,code',[('SEQ_EQ',' SEQ WPSC_EQ \\* ARABIC '),('STYLEREF',' STYLEREF "Heading 1" \\s ')])
def test_reverse_position_snapshot_matches_frozen_writer(monkeypatch,frozen_writer,owner,kind,code):
    baseline=frozen_writer()
    baseline._doc=SimpleNamespace(ComputeStatistics=lambda kind:1)
    baseline._iter_section_page_fields=lambda:[]
    baseline._wpsc_native_fields=[(owner,kind,SimpleNamespace(Result=SimpleNamespace(Text=value)),'numbering') for value in ('FIRST','SECOND')]
    s=MacWordSession()
    s._tracked_numbering=[(NativeIndexHandle('s',name,owner,kind,'numbering'),code) for name in ('first','second')]
    monkeypatch.setattr(s,'_execute',lambda lines:[['stats',1],['identity','first',100],['identity','second',10],
        ['field','story:main text/chain:1',1,kind,code,10,'SECOND',0,0,6],
        ['field','story:main text/chain:1',2,kind,code,100,'FIRST',0,0,5]])
    assert {f.stable_key:f for f in s.snapshot_fields()}=={f.stable_key:f for f in baseline.snapshot_fields()}
    assert [r[4] for r in s._observed_field_topology]==[10,100]


def test_main_ref_identity_ignores_same_offset_header_ref(monkeypatch):
    s=MacWordSession();code=' REF target \\h '
    s._tracked_references=[(NativeIndexHandle('s','own','ref:one','REF','reference'),code)]
    monkeypatch.setattr(s,'_execute',lambda lines:[['stats',1],['identity','own',10],
        ['field','story:main text/chain:1',1,'REF',code,10,'main',0,0,4],
        ['field','story:primary header/chain:1',1,'REF',code,10,'header',0,0,6]])
    assert [f.stable_key for f in s.snapshot_fields()]==[('ref:one','REF',0),('story:primary header/chain:1','REF',0)]


@pytest.mark.parametrize('bad',[None,3,True,{},'field'])
def test_field_malformed_nested_rows_quarantine_after_native_mutation(monkeypatch,bad):
    s,calls,m=bind_cursor(monkeypatch,[[['selection',0,0]],[bad]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.field('SEQ WPSC_EQ \\* ARABIC','n','SEQ_EQ')
    assert s._quarantined and s._structural_changed and not hasattr(s,'_tracked_numbering')


@pytest.mark.parametrize('action', ['bookmark','format'])
def test_numbering_ack_booleans_and_integers_are_not_interchangeable(monkeypatch,action):
    ack=[['bookmark','b',False,True]] if action=='bookmark' else [['format',1,1]]
    s,calls,m=bind_cursor(monkeypatch,[[['selection',0,0]],ack]);c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError):
        if action=='bookmark': c.bookmark('b',0,1)
        else: c.format_paragraph(0)
    assert s._quarantined


@pytest.mark.parametrize('failure',['script-write','deadline','read-only','submitted-error'])
def test_real_transport_distinguishes_presubmission_failure(monkeypatch,tmp_path,failure):
    from pathlib import Path
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=MacWordSession();s.staging_root=tmp_path;s._bound_path='/private/owned.docx';s._owns_doc=True
    reasons=[];calls=[]
    s.lock=SimpleNamespace(quarantine_path=tmp_path/'quarantine.json',quarantine=lambda data:reasons.append(data['reason']))
    s._observed_field_topology=(('before',1),);s._pending_heading=(0,4,1)
    c=implementation().BoundNumberingCursor.__new__(implementation().BoundNumberingCursor)
    c.session=s;c.start=c.position=0
    original_write=Path.write_text
    def write(path,data,**kwargs):
        if 'set content of numberRange' in data:
            if failure=='script-write': raise OSError('local disk full')
            if failure=='deadline': s._deadline=-1
        return original_write(path,data,**kwargs)
    monkeypatch.setattr(Path,'write_text',write)
    def run(command,**kwargs):
        calls.append(command)
        return subprocess.CompletedProcess(command,1,'','execution error: NUMBERING_PARTIAL_FAILURE (-2700)')
    monkeypatch.setattr(transport.subprocess,'run',run)
    if failure=='read-only': s._read_only=True
    with pytest.raises((OSError,NativeWordError,ValueError)): c.text('x')
    if failure=='submitted-error':
        assert len(calls)==1 and s._quarantined and reasons==['Native numbering completion uncertain']
        assert not hasattr(s,'_observed_field_topology') and s._pending_heading is None
    else:
        assert not calls and s._observed_field_topology==( ('before',1), )
        assert s._pending_heading==(0,4,1)
        if failure=='deadline':
            assert s._quarantined and reasons==['Session deadline expired before verified document close']
        else: assert not s._quarantined and not reasons
    assert not getattr(s,'_field_topology_mutation_pending',False)
