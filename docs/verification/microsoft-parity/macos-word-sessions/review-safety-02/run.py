from pathlib import Path
import hashlib,json,shutil,traceback
from skills.WPSComposer import create_document
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
def main():
    out=Path(__file__).resolve().parent
    report={'passed':False,'checks':{}};s=None
    inventory=['set nativeRows to {}','repeat with di from 1 to count documents','set d to document di','if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text,posix full name of d as text,saved of d}','end repeat']
    def texts(snapshot):return [p['text'] for p in snapshot['paragraphs'] if p['text']]
    try:
        with create_document('writer',engine='msoffice') as s:
            before=s._execute(inventory)
            s.apply_format_patch('paragraph:1',text='First\rTerminal')
            assert s._structural_changed
            snapshot=s.inspect_document();assert texts(snapshot)==['First','Terminal']
            count=len(list(s.staging_root.glob('*.applescript')))
            r=s.apply_format_patch('paragraph:1',text='Must not replace',font={'size':-1})
            assert not r['accepted'] and r['rejected']==['font.size']
            assert len(list(s.staging_root.glob('*.applescript')))==count
            report['checks']['invalid_patch_before_native']=True
            s.apply_structural_op({'op':'clone','target':'paragraph:1','to':'end'})
            assert texts(s.inspect_document())==['First','Terminal','First']
            s.apply_structural_op({'op':'move','target':'paragraph:2','to':'end'})
            assert texts(s.inspect_document())==['First','First','Terminal']
            report['checks']['terminal_paragraph_clone_move']=True
            s.save(out/'source.docx');s.export_pdf(out/'terminal.pdf')
            report['checks']['unrelated_unchanged']=before==s._execute(inventory)
            shutil.copytree(s.staging_root,out/'create-runtime')
        original=hashlib.sha256((out/'source.docx').read_bytes()).hexdigest()
        current=out/'current.docx';shutil.copy2(out/'source.docx',current)
        with MacWordSession.open_document(current) as s:
            s.apply_format_patch('paragraph:1',text='Explicit saved')
            assert s.save_current()==str(current)
            s.apply_format_patch('paragraph:1',text='Explicit close saved')
            shutil.copytree(s.staging_root,out/'save-runtime')
            s.close(save_changes=True)
        with MacWordSession.open_document(current,read_only=True) as s:
            snap=s.inspect_document();assert texts(snap)==['Explicit close saved','First','Terminal']
            (out/'reopened.json').write_text(json.dumps(snap,ensure_ascii=False,indent=2))
            s.export_pdf(out/'current.pdf');shutil.copytree(s.staging_root,out/'reopen-runtime')
        shutil.copy2(current,out/'saved-current.docx')
        report['checks']['save_current_and_close_save_reopened']=True
        with MacWordSession.open_document(current) as s:
            shutil.copy2(out/'source.docx',current)
            try:s.close(save_changes=True)
            except ValueError:pass
            else:raise AssertionError('source conflict accepted')
            assert s.lock.file is not None and not s._closed
            assert hashlib.sha256(current.read_bytes()).hexdigest()==original
            shutil.copytree(s.staging_root,out/'conflict-runtime')
            report['checks']['conflict_keeps_lock_until_discard']=True
        assert s._closed and s.lock.file is None
        assert hashlib.sha256((out/'source.docx').read_bytes()).hexdigest()==original
        report['passed']=all(report['checks'].values())
    except BaseException as exc:
        report['error']=str(exc);(out/'failure.txt').write_text(traceback.format_exc())
        if s and s.staging_root and s.staging_root.exists():shutil.copytree(s.staging_root,out/'failure-runtime',dirs_exist_ok=True)
    finally:(out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps(report));raise SystemExit(0 if report['passed'] else 1)

if __name__=="__main__":
    main()
