"""One offline range-representation diagnostic; no native runner or public API."""
from __future__ import annotations
from fixtures.microsoft_parity import macos_word_semantic_table_conversion_probe as previous


def commands(path,name,token):
    lines=previous.commands(path,name,token,end=19)
    old='set diagRange to text object of diagSelection'
    assert lines.count(old)==1
    lines[lines.index(old)]='set diagRange to create range boundDoc start 12 end 19'
    return lines


def valid(rows,path):
    return previous.valid(rows,path,end=19)


def assess(rows):
    """Interpret only a complete, validated ACK; actual dimensions stay explicit."""
    converted=rows[-2]
    result=previous.assess(converted)
    result['table_count_delta']=rows[2][1]-rows[0][4]
    result['returned_dimensions']=converted[4:6] if converted[1] else None
    result['requested_dimensions_observed']=converted[1] and converted[4:6]==[4,3]
    matched=[row for row in rows[3:-2] if row[2:7]==converted[2:7]] if converted[1] else []
    result['document_table_ordinal']=matched[0][1] if len(matched)==1 else None
    return result
