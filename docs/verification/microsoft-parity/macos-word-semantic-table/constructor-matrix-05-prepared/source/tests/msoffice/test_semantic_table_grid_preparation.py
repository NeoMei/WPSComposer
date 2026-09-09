"""Grid shape, exact replacement extent, and placeholder containment gates."""
import importlib
import pytest


def mod():
    try:return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table_grid_probe')
    except ModuleNotFoundError:pytest.fail('Grid preparation is missing')


def test_ascii_grid_has_twelve_nonempty_cells_and_exact_utf16_extent():
    m=mod();rows=m.GRID.split('\r')
    assert len(rows)==4 and all(len(row.split('\t'))==3 for row in rows)
    assert len(m.TOKENS)==len(set(m.TOKENS))==12
    assert all(token and token.isascii() for token in m.TOKENS)
    assert not m.GRID.endswith('\r')
    assert m.GRID_END==12+len(m.GRID.encode('utf-16-le'))//2==131


def test_prepopulate_then_rebind_verify_and_convert_exactly_once():
    m=mod();lines=m.commands('/owned.docx','sentinel','token');joined='\n'.join(lines)
    write=next(i for i,line in enumerate(lines) if line.startswith('set content of diagRange to '))
    rebind=lines.index('set diagRange to create range boundDoc start 12 end 131')
    convert=lines.index('set diagTable to convert to table diagRange separator separate by tabs number of rows 4 number of columns 3')
    assert write<rebind<convert
    assert 'GRID_PREPOPULATION_MISMATCH' in joined and 'GRID_RANGE_CHANGED' in joined
    assert joined.count('convert to table diagRange')==1
    assert 'on error' not in joined and 'set diagResult to "ordinary-error"' not in joined
    assert sum('get cell from table diagTable' in line for line in lines)==12


def ack_rows():
    m=mod();tabletext='\r\x07'.join(m.TOKENS)+'\r\x07'
    base=[['before','/owned.docx',12,19,1,'before REPLACE','EXISTING TABLE 原样','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r',True,'owned',True],
          ['grid',12,m.GRID_END,m.GRID],['attempt','created',0,''],['after',2,'PREFIX '+tabletext+' SUFFIX',False],
          ['table',1,12,148,4,3,tabletext],['table',2,180,199,1,1,'EXISTING TABLE 原样'],
          ['conversion',True,12,148,4,3,tabletext,'PREFIX 中文😀\r','SUFFIX preserved 中文😀\r'],
          ['surroundings','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r','EXISTING TABLE 原样']]
    return base+[['grid-cell',r,c,m.TOKENS[(r-1)*3+c-1]+'\r\x07'] for r in range(1,5) for c in range(1,4)]


def test_complete_grid_ack_contains_all_placeholders_only_in_new_cells():
    assert mod().valid(ack_rows(),'/owned.docx')


@pytest.mark.parametrize('mutation',['wrong_grid_end','missing_cell','wrong_cell_text','outside_placeholder','wrong_dimensions','marker_left_outside'])
def test_incomplete_grid_or_placeholder_leak_cannot_pass(mutation):
    r=ack_rows()
    if mutation=='wrong_grid_end':r[1][2]-=1
    if mutation=='missing_cell':r.pop()
    if mutation=='wrong_cell_text':r[-1][3]='not the expected cell'
    if mutation=='outside_placeholder':r[6][8]+=mod().TOKENS[0]
    if mutation=='wrong_dimensions':r[4][4]=1;r[6][4]=1
    if mutation=='marker_left_outside':r[3][2]+='REPLACE'
    assert not mod().valid(r,'/owned.docx')
