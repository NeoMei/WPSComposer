"""Native Excel owned-session acceptance. Output directory must be new."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import traceback
import xml.etree.ElementTree as ET
import zipfile

from skills.WPSComposer.scripts.msoffice.macos_excel_session import MacExcelSession


def run(source,output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    source=Path(source).resolve();before=hashlib.sha256(source.read_bytes()).hexdigest()
    report={'source':str(source),'source_sha256_before':before,'verified':False,'jobs':[]}
    implementation=Path(__file__).resolve().parents[2]/'skills/WPSComposer/scripts/msoffice/macos_excel_session.py'
    report['source_sha256']=hashlib.sha256(implementation.read_bytes()).hexdigest()
    (output/'session-source.py').write_bytes(implementation.read_bytes())
    session=None
    try:
        session=MacExcelSession.open_document(source)
        report['jobs'].append(str(session._job))
        initial=session.inspect_document(max_cells=100)
        (output/'before.json').write_text(json.dumps(initial,indent=2,ensure_ascii=False)+'\n')
        assert initial['sheet_count']==2
        assert next(c for c in initial['sheets'][1]['cells'] if c['address']=='$B$1')['value']==120
        chart_name=initial['sheets'][0]['shapes'][0]['name']
        report['patches']=[]
        report['patches'].append(session.apply_format_patch('sheet:1/cell:B2',value=15,font={'bold':True,'italic':True,'size':14,'color':'#123456'},fill={'color':'#FFF2CC'},number_format='0.000',horizontal_alignment=-4108,vertical_alignment=-4108,wrap_text=True,indent=1,row_height=30,column_width=20,borders={'7':{'style':1,'weight':2,'color':'#C00000'}}))
        report['patches'].append(session.apply_format_patch('sheet:1/chart:1',chart_title='Updated native revenue',chart_type=51,geometry={'left':30,'width':420}))
        report['patches'].append(session.apply_format_patch('sheet:1/shape:@name='+chart_name,name='StableRevenue'))
        report['patches'].append(session.apply_format_patch('sheet:2',name='Results',page_setup={'orientation':2,'top_margin':40,'print_area':'$A$1:$D$4'}))
        report['structural']=[]
        for op in [
            {'op':'insert','parent':'sheet:1','type':'row','position':{'index':10},'props':{'values':['Inserted',99]}},
            {'op':'remove','target':'sheet:1/cell:A10','axis':'row'},
            {'op':'insert','parent':'sheet:1','type':'column','position':{'index':5},'props':{'values':['Inserted column',88]}},
            {'op':'remove','target':'sheet:1/cell:E1','axis':'column'},
            {'op':'insert','type':'sheet','props':{'name':'Temporary'}},
            {'op':'remove','target':'sheet:3'},
        ]:
            report['structural'].append(session.apply_structural_op(op))
            if op.get('op')=='insert' and op.get('type') in ('row','column'):
                inserted=session.inspect_document(max_cells=100)
                address='$A$10' if op['type']=='row' else '$E$1'
                cell=next(c for c in inserted['sheets'][0]['cells'] if c['address']==address)
                assert cell['value']==('Inserted' if op['type']=='row' else 'Inserted column')
        # Reinspect after structural edits before using a positional cell target.
        session.apply_format_patch('sheet:1/shape:@name=StableRevenue',geometry={'left':35},fill={'color':'#F0F0F0','transparency':0.1},line={'color':'#123456','weight':2,'visible':True})
        final=session.inspect_document(max_cells=100)
        (output/'after.json').write_text(json.dumps(final,indent=2,ensure_ascii=False)+'\n')
        b2=next(c for c in final['sheets'][0]['cells'] if c['address']=='$B$2')
        assert b2['value']==15 and b2['font']['bold'] and b2['font']['italic'] and b2['font']['size']==14
        assert b2['font']['color']=='#123456' and b2['fill']['color']=='#FFF2CC'
        assert final['sheets'][1]['name']=='Results'
        assert final['sheets'][0]['shapes'][0]['name']=='StableRevenue'
        assert final['sheets'][0]['charts'][0]['title']=='Updated native revenue'
        report['selection']=session.inspect_selection()
        assert report['selection']['id']=='selection'
        session.save(output/'edited.xlsx');session.export_pdf(output/'edited.pdf')
        session.close();session=None
        assert hashlib.sha256(source.read_bytes()).hexdigest()==before
        with MacExcelSession.open_document(output/'edited.xlsx',read_only=True) as reopened:
            report['jobs'].append(str(reopened._job))
            snap=reopened.inspect_document(max_cells=100)
            (output/'reopened.json').write_text(json.dumps(snap,indent=2,ensure_ascii=False)+'\n')
            assert next(c for c in snap['sheets'][0]['cells'] if c['address']=='$B$2')['value']==15
            assert next(c for c in snap['sheets'][0]['cells'] if c['address']=='$B$6')['value']==65
            assert next(c for c in snap['sheets'][1]['cells'] if c['address']=='$B$1')['value']==130
        # Atomic-style abort: earlier staged mutation must never touch source or publish.
        with MacExcelSession.open_document(source) as aborted:
            report['jobs'].append(str(aborted._job))
            aborted.apply_format_patch('sheet:1/cell:B2',value=999)
            try: aborted.apply_format_patch('sheet:1/cell:B3',unknown='invalid')
            except ValueError: pass
            else: raise AssertionError('Invalid patch was accepted')
        assert not (output/'aborted.xlsx').exists()
        assert hashlib.sha256(source.read_bytes()).hexdigest()==before
        import fitz
        pdf=fitz.open(output/'edited.pdf');text='\n'.join(p.get_text() for p in pdf)
        assert len(pdf)==2 and 'Updated native revenue' in text and '130' in text
        for i,page in enumerate(pdf): page.get_pixmap(matrix=fitz.Matrix(1.25,1.25)).save(output/f'page-{i+1}.png')
        (output/'pdf-text.txt').write_text(text)
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(output/'edited.xlsx') as archive:
            styles=ET.fromstring(archive.read('xl/styles.xml'));sheet=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            cell=sheet.find(".//s:c[@r='B2']",ns);xf=styles.find('s:cellXfs',ns)[int(cell.get('s'))]
            border=styles.find('s:borders',ns)[int(xf.get('borderId'))]
            assert border.find('s:left',ns).get('style')=='thin'
            assert border.find('s:left/s:color',ns).get('rgb')=='FFC00000'
        report.update(verified=True,pdf_pages=len(pdf),atomic_abort_source_preserved=True,native_border_verified=True)
    except BaseException as error:
        report['error']=type(error).__name__+': '+str(error)
        (output/'error.log').write_text(traceback.format_exc())
    finally:
        if session is not None: session.close()
        report['source_sha256_after']=hashlib.sha256(source.read_bytes()).hexdigest()
        for index,job in enumerate(report['jobs']):
            logs=output/f'job-{index+1}-logs';logs.mkdir()
            for file in Path(job).glob('step-*'):shutil.copyfile(file,logs/file.name)
            if (Path(job)/'recovery.json').exists():shutil.copyfile(Path(job)/'recovery.json',logs/'recovery.json')
        report['checksums']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in output.iterdir() if p.is_file()}
        (output/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n')
    return report


def run_active(source,output):
    import tempfile
    from skills.WPSComposer.scripts.msoffice.macos_excel_session import _quote
    from skills.WPSComposer.scripts.msoffice.macos_office_runtime import _container_root
    from skills.WPSComposer.scripts.msoffice.input_validation import validate_native_input
    from skills.WPSComposer.scripts.document_api import inspect,edit
    out=Path(output).resolve();out.mkdir(parents=True,exist_ok=False)
    source=Path(source).resolve()
    validate_native_input(source,'spreadsheet')
    job=Path(tempfile.mkdtemp(prefix='session-active-fixture-',dir=_container_root('spreadsheet')))
    native=job/'active-owned.xlsx';shutil.copyfile(source,native)
    report={'verified':False,'source':str(source),'source_sha256_before':hashlib.sha256(source.read_bytes()).hexdigest(),'native':str(native),'jobs':[]}
    def run(body,label):
     script='with timeout of 60 seconds\ntell application "/Applications/Microsoft Excel.app"\n'+body+'\nend tell\nend timeout'
     (out/(label+'.applescript')).write_text(script)
     r=subprocess.run(['/usr/bin/osascript','-e',script],text=True,capture_output=True,timeout=60)
     (out/(label+'.stdout.log')).write_text(r.stdout);(out/(label+'.stderr.log')).write_text(r.stderr)
     assert r.returncode==0,r.stderr
     return r.stdout.strip()
    try:
     run('set b to open workbook workbook file name '+_quote(str(native))+'\nactivate object worksheet "Data" of b\nselect range "B2" of worksheet "Data" of b\nreturn full name of b','open')
     with MacExcelSession.attach_active() as session:
      report['jobs'].append(str(session._job));assert session.is_bound_to(native)
      before=session.inspect_selection();assert before['address']=='$B$2' and before['value']==10
      snap=session.inspect_document(max_cells=100);after=session.inspect_selection();assert before==after
      session.preflight_save()
      try:session.preflight_save(out/'unsupported-copy.xlsx')
      except NotImplementedError:report['save_copy_preflight_rejected']=True
      else:raise AssertionError('attached copy should fail before mutation')
      report['patch']=session.apply_format_patch('selection',value=17)
      session.save_current()
     report['stays_open']=run('set b to workbook "active-owned.xlsx"\nif full name of b is not '+_quote(str(native))+' then error "identity"\nreturn value of range "B2" of worksheet "Data" of b','after-close')=='17.0'
     assert report['stays_open']
     report['public_edit']=edit(None,kind='sheet',engine='msoffice',patches=[{'target':'sheet:1/cell:B2','value':19}],inspect_after=True)
     assert report['public_edit']['ok']
     report['public_active_inspect']=inspect(None,kind='sheet',engine='msoffice',max_cells=100)
     run('set b to workbook "active-owned.xlsx"\nif full name of b is not '+_quote(str(native))+' then error "identity"\nclose b saving no\nreturn "closed owned"','close-owned')
     with MacExcelSession.open_document(native,read_only=True) as reopened:
      report['jobs'].append(str(reopened._job));snap=reopened.inspect_document(max_cells=100)
      assert next(c for c in snap['sheets'][0]['cells'] if c['address']=='$B$2')['value']==19
     report['source_sha256_after']=hashlib.sha256(source.read_bytes()).hexdigest();assert report['source_sha256_before']==report['source_sha256_after']
     shutil.copyfile(native,out/'active-edited.xlsx');report['verified']=True
    except BaseException as error:
     report['error']=str(error);(out/'error.log').write_text(traceback.format_exc())
    finally:
     for i,j in enumerate(report['jobs']):shutil.copytree(j,out/f'native-logs-{i+1}')
     p=Path('skills/WPSComposer/scripts/msoffice/macos_excel_session.py');shutil.copyfile(p,out/'session-source.py');report['session_source_sha256']=hashlib.sha256(p.read_bytes()).hexdigest()
     (out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')

    return report


def run_business(source,output):
    source=Path(source).resolve();output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'verified':False,'source':str(source),'source_sha256_before':hashlib.sha256(source.read_bytes()).hexdigest(),'jobs':[]}
    implementation=Path(__file__).resolve().parents[2]/'skills/WPSComposer/scripts/msoffice/macos_excel_session.py'
    (output/'session-source.py').write_bytes(implementation.read_bytes());report['session_source_sha256']=hashlib.sha256(implementation.read_bytes()).hexdigest()
    session=None
    try:
        session=MacExcelSession.open_document(source);report['jobs'].append(str(session._job))
        session.select_sheet(1)
        session.write_table(1,1,[['Metric','Value'],['Alpha',10],['Beta',20],['Total','=SUM(B2:B3)']])
        session.write_cell(2,2,15);session.set_formula(4,2,'=SUM(B2:B3)')
        session.set_cell_style(2,2,bold=True,italic=True,font_size=13,font_color=0x563412,fill_color='#FFF2CC',align=-4108,number_format='0.00')
        session.set_range_style('A3:B3',bold=True,font_color='#008000')
        session.set_borders('A1:B4',color='#123456')
        session.add_title_row('A6:C6','Native business title')
        session.merge_cells('A8:C8');session.write_cell(8,1,'Merged native row')
        report['freeze']=session.freeze_panes('B2')
        assert report['freeze']=={'freeze_panes':True,'split_column':1,'split_row':1}
        session.set_column_width(1,25);session.set_row_height(1,28);session.autofit()
        report['chart']=session.add_chart(chart_type=51,left=20,top=180,width=420,height=240,source_range='A1:B3',title='Native revenue business')
        session.conditional_format('B2:B3',formula='0')
        report['conditional']=session.conditional_format('B2:B3',formula='15')
        assert report['conditional']['condition_count']==1
        session.set_header_footer(left='NATIVE HEADER',center='&P',right='NATIVE RIGHT')
        session.rename_sheet(2,'Results');session.select_sheet(2);session.set_range_style('A1:B1',bold=True)
        new=session.add_sheet('Business report');session.write_cell(1,1,'Native third sheet')
        assert new['path']=='sheet:3'
        session.select_sheet(1)
        try:session.apply_structural_op({'op':'remove','target':'sheet:2'})
        except NotImplementedError:report['existing_sheet_delete_preflight_rejected']=True
        else:raise AssertionError('existing worksheet deletion must reject without mutating')
        snapshot=session.inspect_document(max_cells=100)
        (output/'snapshot.json').write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n')
        assert [s['name'] for s in snapshot['sheets']]==['Data','Results','Business report']
        assert next(c for c in snapshot['sheets'][0]['cells'] if c['address']=='$B$4')['value']==35
        session.save_xlsx(output/'business.xlsx');session.export_pdf(output/'business.pdf');session.close();session=None
        with MacExcelSession.open_document(output/'business.xlsx',read_only=True) as reopened:
            report['jobs'].append(str(reopened._job));snapshot=reopened.inspect_document(max_cells=100)
            assert snapshot['sheet_count']==3 and snapshot['sheets'][0]['charts'][0]['title']=='Native revenue business'
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        with zipfile.ZipFile(output/'business.xlsx') as archive:
            root=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            pane=root.find('.//s:pane',ns);assert pane.get('state') in ('frozen','frozenSplit') and pane.get('xSplit')=='1' and pane.get('ySplit')=='1'
            rule=root.find('.//s:cfRule',ns);assert rule.get('type')=='cellIs' and rule.get('operator')=='greaterThan' and rule.find('s:formula',ns).text=='15'
            assert root.find('.//s:oddHeader',ns).text=='&LNATIVE HEADER&CP&RNATIVE RIGHT' or 'NATIVE HEADER' in root.find('.//s:oddHeader',ns).text
            assert root.find('.//s:mergeCell[@ref="A6:C6"]',ns) is not None
            chart=ET.fromstring(archive.read('xl/charts/chart1.xml'));assert any(n.text=='Native revenue business' for n in chart.iter())
        import fitz
        pdf=fitz.open(output/'business.pdf');report['pdf_pages']=len(pdf)
        text='\n'.join(p.get_text() for p in pdf);assert 'NATIVE HEADER' in text and 'Native revenue business' in text and 'Native third sheet' in text
        (output/'pdf-text.txt').write_text(text)
        for i,page in enumerate(pdf):page.get_pixmap(matrix=fitz.Matrix(1.25,1.25)).save(output/f'page-{i+1}.png')
        report['source_sha256_after']=hashlib.sha256(source.read_bytes()).hexdigest();assert report['source_sha256_before']==report['source_sha256_after']
        report['verified']=True
    except BaseException as e:report['error']=str(e);(output/'error.log').write_text(traceback.format_exc())
    finally:
        if session is not None:session.close()
        for i,job in enumerate(report['jobs']):shutil.copytree(job,output/f'native-logs-{i+1}')
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    return report


def run_structure(source,output):
    source=Path(source).resolve();output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'verified':False,'source_sha256_before':hashlib.sha256(source.read_bytes()).hexdigest(),'cases':[],'jobs':[]}
    implementation=Path(__file__).resolve().parents[2]/'skills/WPSComposer/scripts/msoffice/macos_excel_session.py'
    (output/'session-source.py').write_bytes(implementation.read_bytes());report['session_source_sha256']=hashlib.sha256(implementation.read_bytes()).hexdigest()
    try:
        for axis in ('row','column'):
            for verb in ('clone','move','remove'):
                with MacExcelSession.open_document(source) as session:
                    report['jobs'].append(str(session._job))
                    session.write_table(1,1,[['C1','C2','C3','C4','C5']]+[[f'R{r}',r*10+2,r*10+3,r*10+4,r*10+5] for r in range(2,9)])
                    result=session.apply_structural_op({'op':verb,'target':'sheet:1/range:B2:D4','axis':axis,'to':{'index':7 if axis=='row' else 5}})
                    snapshot=session.inspect_document(max_cells=100)
                    cells={c['address'].replace('$',''):c['value'] for c in snapshot['sheets'][0]['cells']}
                    if axis=='row':
                        if verb=='clone':assert cells['A3']=='R3' and cells['A4']=='R4' and cells['A7']=='R2' and cells['A8']=='R7' and cells['A9']=='R8'
                        elif verb=='move':assert cells['A2']=='R3' and cells['A3']=='R4' and cells['A6']=='R2' and cells['A7']=='R7'
                        else:assert cells['A2']=='R3' and cells['A3']=='R4' and cells['A4']=='R5'
                    else:
                        if verb=='clone':assert cells['C1']=='C3' and cells['D1']=='C4' and cells['E1']=='C2' and cells['F1']=='C5'
                        elif verb=='move':assert cells['B1']=='C3' and cells['C1']=='C4' and cells['D1']=='C2' and cells['E1']=='C5'
                        else:assert cells['B1']=='C3' and cells['C1']=='C4' and cells['D1']=='C5'
                    report['cases'].append({'axis':axis,'verb':verb,'result':result,'cells':cells})
        with MacExcelSession.open_document(source) as session:
            report['jobs'].append(str(session._job));session.select_sheet(2)
            report['sheet_move']=session.apply_structural_op({'op':'move','target':'sheet:2','to':{'before':1}})
            session.write_cell(10,1,'Selected original Summary')
            report['sheet_clone']=session.apply_structural_op({'op':'clone','target':'sheet:1','to':{'after':2}})
            session.write_cell(11,1,'Selection retained after clone')
            snap=session.inspect_document(max_cells=100)
            assert [s['name'] for s in snap['sheets']]==['Summary','Data','Summary (2)']
            cells={c['address']:c['value'] for c in snap['sheets'][0]['cells']}
            assert cells['$A$10']=='Selected original Summary' and cells['$A$11']=='Selection retained after clone'
            session.save(output/'structural.xlsx');session.export_pdf(output/'structural.pdf')
        with MacExcelSession.open_document(output/'structural.xlsx',read_only=True) as reopened:
            report['jobs'].append(str(reopened._job));assert reopened.inspect_document(max_cells=100)['sheet_count']==3
        report['source_sha256_after']=hashlib.sha256(source.read_bytes()).hexdigest();assert report['source_sha256_before']==report['source_sha256_after']
        report['verified']=True
    except BaseException as e:report['error']=str(e);(output/'error.log').write_text(traceback.format_exc())
    finally:
        for i,job in enumerate(report['jobs']):shutil.copytree(job,output/f'native-logs-{i+1}')
        (output/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n')
    return report


def run_new(source,output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'verified':False,'jobs':[]}
    implementation=Path(__file__).resolve().parents[2]/'skills/WPSComposer/scripts/msoffice/macos_excel_session.py'
    (output/'session-source.py').write_bytes(implementation.read_bytes());report['session_source_sha256']=hashlib.sha256(implementation.read_bytes()).hexdigest()
    try:
        from skills.WPSComposer import create_document
        with create_document('sheet',engine='msoffice') as session:
            report['jobs'].append(str(session._job));assert session._native.is_file()
            initial=session.inspect_document(max_cells=100);assert initial['sheet_count']>=1 and all(not s['cells'] for s in initial['sheets'])
            session.write_cell(1,1,'clear me');session.write_cell(1,1,None)
            assert not session.inspect_document(max_cells=100)['sheets'][0]['cells']
            session.rename_sheet(1,'Native created');session.add_title_row('A1:D1','Native new workbook')
            session.write_table(3,1,[['Metric','Value'],['New',12],['Doubled','=B4*2']])
            session.freeze_panes('A4')
            session.add_chart(source_range='A3:B5',title='Native-created chart')
            session.save_xlsx(output/'new.xlsx');session.export_pdf(output/'new.pdf')
        with MacExcelSession.open_document(output/'new.xlsx',read_only=True) as reopened:
            report['jobs'].append(str(reopened._job));snapshot=reopened.inspect_document(max_cells=100)
            assert snapshot['sheets'][0]['name']=='Native created'
            assert snapshot['sheets'][0]['freeze_panes'] is True
            assert next(c for c in snapshot['sheets'][0]['cells'] if c['address']=='$B$5')['value']==24
            (output/'reopened.json').write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n')
        report['verified']=True
    except BaseException as e:report['error']=str(e);(output/'error.log').write_text(traceback.format_exc())
    finally:
        for i,job in enumerate(report['jobs']):shutil.copytree(job,output/f'native-logs-{i+1}')
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--source',required=True,type=Path);parser.add_argument('--output-dir',required=True,type=Path)
    parser.add_argument('--mode',choices=('full','active','business','structural','new'),default='full')
    args=parser.parse_args();result=({'full':run,'active':run_active,'business':run_business,'structural':run_structure,'new':run_new}[args.mode])(args.source,args.output_dir);print(json.dumps(result,indent=2,ensure_ascii=False));raise SystemExit(0 if result['verified'] else 1)
