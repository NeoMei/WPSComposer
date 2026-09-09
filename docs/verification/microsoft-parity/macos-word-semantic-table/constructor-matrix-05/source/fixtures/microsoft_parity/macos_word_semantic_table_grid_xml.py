"""Grid-only XML gate with the reviewed semantic fixture's exact outside splice.

The three nested preservation functions are copied verbatim from the frozen
semantic fixture; candidate discovery and cell topology are diagnostic-specific.
"""
from __future__ import annotations
import xml.etree.ElementTree as ET
from fixtures.microsoft_parity.macos_word_semantic_table_grid_probe import TOKENS
from fixtures.microsoft_parity.macos_word_semantic_table import W,OLD


def checks(before,after):
    old,new=ET.fromstring(before),ET.fromstring(after)
    text=lambda n: ''.join(e.text or '' for e in n.iter(W+'t'))
    tables=list(new.iter(W+'tbl'));candidates=[t for t in tables if TOKENS[0] in text(t)]
    result={'grid_4x3':False,'cells_exact':False,'outside_preserved':False}
    if len(candidates)!=1:return result
    table=candidates[0];rows=table.findall(W+'tr');cells=[r.findall(W+'tc') for r in rows]
    result['grid_4x3']=(len(tables)==2 and len(rows)==4 and all(len(r)==3 for r in cells)
        and len(table.findall('./'+W+'tblGrid/'+W+'gridCol'))==3)
    result['cells_exact']=[text(c) for row in cells for c in row]==list(TOKENS)
    def normalized(node):
        return (node.tag,tuple(sorted((k,v) for k,v in node.attrib.items() if not k.split('}')[-1].startswith('rsid')
            and k.split('}')[-1] not in ('paraId','textId'))),node.text or '',tuple(normalized(c) for c in node
            if c.tag not in (W+'bookmarkStart',W+'bookmarkEnd')))
    def terminal_paragraph(original, tail):
        # The one allowed tail is an empty paragraph with the replacement
        # paragraph's exact properties. Empty unformatted runs are equivalent
        # to no run; breaks, fields, section/page properties or other content
        # are never discarded merely because there is no w:t text.
        if tail.tag != W+'p' or normalized(original)[1] != normalized(tail)[1]:
            return False
        def properties(paragraph):
            found=paragraph.findall(W+'pPr')
            if len(found)>1: return ('invalid-duplicate-properties',)
            return normalized(found[0]) if found else None
        if properties(original) != properties(tail): return False
        for child in original:
            if child.tag in (W+'pPr',W+'bookmarkStart',W+'bookmarkEnd'): continue
            if child.tag != W+'r' or any(part.tag not in (W+'rPr',W+'t') for part in child):
                return False
        for child in tail:
            if child.tag in (W+'pPr',W+'bookmarkStart',W+'bookmarkEnd'): continue
            if child.tag != W+'r' or normalized(child)[1]: return False
            for part in child:
                if part.tag != W+'t' or part.text or len(part): return False
                if any(key != '{http://www.w3.org/XML/1998/namespace}space' for key in part.attrib):
                    return False
        return True
    def exact_splice():
        old_body,new_body=old.find(W+'body'),new.find(W+'body')
        if old_body is None or new_body is None: return False
        original,current=list(old_body),list(new_body)
        slots=[i for i,node in enumerate(original) if node.tag==W+'p' and text(node)=='REPLACE']
        if len(slots)!=1 or len(current)!=len(original)+1: return False
        index=slots[0]
        if current[index] is not table or not terminal_paragraph(original[index],current[index+1]):
            return False
        # Preserve every other body node in order, including all existing
        # empty paragraphs and their layout/section/run-level structure.
        before=original[:index]+original[index+1:]
        after=current[:index]+current[index+2:]
        return (normalized(old_body)[1:3]==normalized(new_body)[1:3]
                and [normalized(node) for node in before]==[normalized(node) for node in after])
    result['outside_preserved']=(exact_splice()
        and text(new).count('PREFIX')==text(new).count('SUFFIX')==text(new).count(OLD)==1
        and 'REPLACE' not in text(new) and 'OLD TABLE SLOT' not in text(new))
    return result
