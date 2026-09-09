"""Offline designed 4x3 range-to-grid constructor proposal, never a public API.

Prepopulation is a real mutation. Every later error must propagate to quarantine;
there is deliberately no ordinary-error catch or attempted rollback here.
"""
from __future__ import annotations
from fixtures.microsoft_parity import macos_word_semantic_table_explicit_range_probe as previous
from fixtures.microsoft_parity.macos_word_quality_feasibility import exact,bound_guard,sentinel_guard
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string

TOKENS=tuple(f'WPSC_R{r}C{c}' for r in range(1,5) for c in range(1,4))
GRID='\r'.join('\t'.join(TOKENS[r*3:r*3+3]) for r in range(4))
GRID_END=12+len(GRID.encode('utf-16-le'))//2


def commands(path,name,token):
    lines=previous.commands(path,name,token)
    start=lines.index('try');stop=lines.index('end try')+1
    assert lines[start+1]=='set diagTable to convert to table diagRange number of rows 4 number of columns 3'
    lines[start:stop]=[
        f'set content of diagRange to {apple_string(GRID)}',
        f'set diagRange to create range boundDoc start 12 end {GRID_END}',
        f'if start of content of diagRange is not 12 or end of content of diagRange is not {GRID_END} then error "GRID_RANGE_CHANGED"',
        'set gridBeforeText to content of diagRange as text',
        f'if not {exact("gridBeforeText",GRID)} then error "GRID_PREPOPULATION_MISMATCH"',
        *bound_guard(path),*sentinel_guard(name,token),
        f'set end of nativeRows to {{"grid",12,{GRID_END},gridBeforeText}}',
        'set diagTable to convert to table diagRange separator separate by tabs number of rows 4 number of columns 3',
    ]
    for r in range(1,5):
        for c in range(1,4):
            lines += [f'set gridCell to get cell from table diagTable row {r} column {c}',
                      f'set end of nativeRows to {{"grid-cell",{r},{c},content of text object of gridCell as text}}']
    return lines


def valid(rows,path):
    if not isinstance(rows,list) or len(rows)<20 or rows[1]!=['grid',12,GRID_END,GRID]:return False
    core=rows[:1]+rows[2:-12]
    if not previous.valid(core,path):return False
    converted=core[-2]
    if not converted[1] or converted[2]!=12 or converted[4:6]!=[4,3] or core[2][1]-core[0][4]!=1:return False
    if 'REPLACE' in core[2][2] or any(token in converted[7] or token in converted[8] for token in TOKENS):return False
    for index,row in enumerate(rows[-12:]):
        r,c=divmod(index,3)
        if not isinstance(row,list) or len(row)!=4 or row[:3]!=['grid-cell',r+1,c+1]:return False
        if type(row[1]) is not int or type(row[2]) is not int or not isinstance(row[3],str):return False
        if row[3].rstrip('\r\x07')!=TOKENS[index]:return False
    return True


def assess(rows):
    result=previous.assess(rows[:1]+rows[2:-12])
    result['prepopulated_range']=[12,GRID_END]
    result['placeholder_cells_verified']=12
    result['placeholders_outside_new_table']=False
    return result
