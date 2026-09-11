"""Offline two-case proposal; no native runner or public capability.

Reuse Matrix02 guards and ACK rows, replacing the constructor with Word's
explicit text-range conversion. Noncollapsed marker text may become cell text.
"""
from __future__ import annotations
from fixtures.microsoft_parity import macos_word_semantic_table_matrix02 as base

CASES=(('A-convert-noncollapsed',19),('B-convert-collapsed',12))


def commands(path,name,token,*,end):
    if type(end) is not int or end not in (12,19):raise ValueError('Closed conversion ranges only')
    lines=base.commands(path,name,token,4,3,19,True,location='diagRange')
    lines=[line.replace('set selection end of selection of boundWindow to 19',f'set selection end of selection of boundWindow to {end}')
           .replace('end of content of diagRange is not 19',f'end of content of diagRange is not {end}')
           for line in lines]
    old='set diagTable to make new table at diagRange with properties {text object:diagRange,number of rows:4,number of columns:3}'
    assert lines.count(old)==1
    lines[lines.index(old)]='set diagTable to convert to table diagRange number of rows 4 number of columns 3'
    return lines+[
        'if diagResult is "created" then',
        'set conversionStart to start of content of text object of diagTable',
        'set conversionEnd to end of content of text object of diagTable',
        'set conversionBefore to create range boundDoc start 0 end conversionStart',
        'set conversionAfter to create range boundDoc start conversionEnd end (end of content of text object of boundDoc)',
        'set end of nativeRows to {"conversion",true,conversionStart,conversionEnd,number of rows of diagTable,number of columns of diagTable,'
        'content of text object of diagTable as text,content of conversionBefore as text,content of conversionAfter as text}',
        'else',
        'set end of nativeRows to {"conversion",false,-1,-1,0,0,"","",""}',
        'end if',
        'set conversionPrefix to create range boundDoc start 0 end 12',
        'set end of nativeRows to {"surroundings",content of conversionPrefix as text,'
        'content of text object of bookmark "semantic_suffix" of boundDoc as text,'
        'content of text object of bookmark "semantic_old_table" of boundDoc as text}',
    ]


def valid(rows,path,*,end):
    if not isinstance(rows,list) or len(rows)<6 or not base.valid(rows[:-2],path,end,True):return False
    converted,surroundings=rows[-2:]
    if not isinstance(converted,list) or len(converted)!=9 or converted[0]!='conversion' or type(converted[1]) is not bool:return False
    if any(type(v) is not int for v in converted[2:6]) or not all(isinstance(v,str) for v in converted[6:]):return False
    if converted[1]:
        if not 0<=converted[2]<converted[3] or not all(v>0 for v in converted[4:6]):return False
        if rows[1][1:3]!=['created',0]:return False
    elif converted!=['conversion',False,-1,-1,0,0,'','',''] or rows[1][1]!='ordinary-error':return False
    if not isinstance(surroundings,list) or len(surroundings)!=4 or surroundings[0]!='surroundings' or not all(isinstance(v,str) for v in surroundings[1:]):return False
    if surroundings[1:]!=[rows[0][7],rows[0][8],rows[0][6]]:return False
    old=[r for r in rows[3:-2] if r[4:6]==[1,1] and r[6]==rows[0][6]]
    if len(old)!=1:return False
    if converted[1]:
        matched=[r for r in rows[3:-2] if r[2:7]==converted[2:7]]
        if len(matched)!=1 or matched[0] is old[0]:return False
    return True


def assess(converted):
    created=converted[1]
    return {'conversion_succeeded':created,'new_table_bounds':converted[2:4] if created else None,
            'starts_at_requested_range':created and converted[2]==12,
            'marker_in_new_table':created and 'REPLACE' in converted[6],
            'marker_outside_new_table':('REPLACE' in converted[7] or 'REPLACE' in converted[8]) if created else None}
