"""Strict, file-only reassessment; preserves the original native FAIL report."""
from pathlib import Path
import hashlib,json
root=Path(__file__).resolve().parent
r=json.loads((root/'report.json').read_text())
def same(a,b):
    return type(a) is type(b) and (len(a)==len(b) and all(same(x,y) for x,y in zip(a,b)) if isinstance(a,list) else a==b)
steps=r['steps'];full=[s for s in steps if s['label'].startswith('production-first-format')]
assert len(full)==2 and all(s['rows'][0][0]=='read-ok' for s in steps)
first=full[0]['rows'][0][1];assert len(first)==220
assert same(first,full[1]['rows'][0][1])
expected=['Arial',17,True,True,'underline single',[4369,8738,13107],[61166,59624,43690]]
assert any(same(row,expected) for row in first)
seen=[]
for step in steps:
    assert step['before'][1:3]==[True,True] and step['after'][1:3]==[True,True]
    assert same(step['before'][3:],r['baseline'][3:]) and same(step['after'][3:],r['baseline'][3:])
    if step['label'].startswith('record-vs-legacy'):
        for tag,position,equal,legacy in step['rows'][0][1]:
            assert tag=='same-coordinate' and type(position) is int and equal is True
            assert same(legacy,first[position]);seen.append(position)
assert seen==list(range(220))
assert r['owned_closed'] is True and r['inventory_final']==[]
assert same(r['after'],r['baseline'])
assert not any(r.get(k) for k in ('blocker','error','first_dirty_phase','remaining_owned_path'))
assert all(r['checks'].get(k) is True for k in ('all_coordinates_equal','full_format_repeat_equal','production_matches_legacy','body_unchanged','private_file_unchanged','inventory_empty','sources_unchanged'))
result={'status':'PASS_SCOPED_CHARACTER_READ','original_native_status':r['status'],
'original_report_sha256':hashlib.sha256((root/'report.json').read_bytes()).hexdigest(),
'correction':'Expected RGB channels must use native 16-bit values, not 8-bit. Production values unchanged.',
'exact_coordinate_count':len(seen),'strict_typed_equality':True,'saved_body_and_measured_counts_preserved':True,
'full_character_read_seconds':[s['seconds'] for s in full],'owned_closed':True,'inventory_empty':True,
'full_quality_snapshot_accepted':False}
(root/'assessment.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result,indent=2))
