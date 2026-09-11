"""Owned macOS Excel file sessions with native inspection and bounded mutation.

Files are copied into Excel's container. A held per-app lock spans the session;
all AppleEvents bind the exact owned path. Failed/uncertain steps quarantine the
app and retain staging. Saved active documents bind their exact native path;
attached close leaves them open. Explicit row/column move/clone uses native
copy/cut and reports its clipboard side effect.
"""
from __future__ import annotations

import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time
from uuid import uuid4

from ..artifact_transport import publish_artifact, copy_file_before_deadline, snapshot_artifact_state, validate_office_package, validate_pdf, validate_before_deadline, ValidatorSpec
from .input_validation import validate_native_input, validate_spreadsheet_formula
from .macos_excel_script import _quote, _rgb, _column_name, _column_number, _sheet_name
from .macos_office_runtime import OfficeJobLock, _container_root

_APP = '/Applications/Microsoft Excel.app'
_JSON = '''use framework "Foundation"
use scripting additions
on j(v)
 if v is missing value then return "null"
 try
  set box to current application's NSArray's arrayWithObject:v
  set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
  set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
  return text 2 thru -2 of s
 on error
  try
   return my j(v as text)
  on error
   return "null"
  end try
 end try
end j
on joined(itemsList)
 set oldDelimiters to AppleScript's text item delimiters
 set AppleScript's text item delimiters to ","
 set resultText to itemsList as text
 set AppleScript's text item delimiters to oldDelimiters
 return resultText
end joined
'''
_H_ALIGN = {1:'horizontal align general', -4108:'horizontal align center', -4131:'horizontal align left', -4152:'horizontal align right', -4130:'horizontal align justify', 5:'horizontal align fill',7:'horizontal align center across selection',-4117:'horizontal align distributed'}
_V_ALIGN = {-4107:'vertical alignment bottom',-4108:'vertical alignment center',-4117:'vertical alignment distributed',-4130:'vertical alignment justify',-4160:'vertical alignment top'}
_UNDERLINE = {-4142:'underline style none',2:'underline style single',-4119:'underline style double',4:'underline style single accounting',5:'underline style double accounting'}
_BORDER = {5:'diagonal down',6:'diagonal up',7:'edge left',8:'edge top',9:'edge bottom',10:'edge right',11:'inside vertical',12:'inside horizontal'}
_BORDER_STYLE={1:'continuous',-4115:'dash',4:'dash dot',5:'dash dot dot',-4118:'dot',-4119:'double',13:'slant dash dot',-4142:'line style none'}
_BORDER_WEIGHT={1:'border weight hairline',2:'border weight thin',-4138:'border weight medium',4:'border weight thick'}
_CHART={51:'column clustered',52:'column stacked',53:'column stacked 100',57:'bar clustered',58:'bar stacked',59:'bar stacked 100',4:'line chart',5:'pie chart',1:'area chart',-4169:'xyscatter',-4120:'doughnut'}


def parse_target(target):
    if target == 'selection': return (None, 'selection', None)
    if not isinstance(target,str): raise ValueError('Invalid Excel target')
    match=re.fullmatch(r'sheet:([1-9]\d*)(?:/(cell|range|shape|chart):(.+))?',target)
    if not match: raise ValueError('Invalid Excel target')
    index=int(match[1]);kind=match[2] or 'sheet';ref=match[3]
    if kind in ('cell','range'):
        address=(ref or '').replace('$','').upper()
        found=re.fullmatch(r'([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?',address)
        if not found or (kind=='cell' and found[3]): raise ValueError('Invalid Excel range')
        left,top=_column_number(found[1]),int(found[2]);right,bottom=_column_number(found[3] or found[1]),int(found[4] or found[2])
        if not (1<=left<=right<=16384 and 1<=top<=bottom<=1048576): raise ValueError('Excel range exceeds bounds')
        return index,kind,address
    if kind in ('shape','chart'):
        if kind=='shape' and ref.startswith('@name='):
            _quote(ref[6:]);return index,'shape_name',ref[6:]
        if kind=='shape' and ref.startswith('@id='): raise NotImplementedError('Mac Excel has no verified stable numeric shape identity')
        if not ref.isdigit() or int(ref)<1: raise ValueError('Invalid Excel object index')
        return index,kind,int(ref)
    return index,kind,None


def _object(fields):
    pieces=[]
    for key,expr in fields.items(): pieces += [_quote(json.dumps(key)+':'),f'my j({expr})'] if not pieces else [_quote(','+json.dumps(key)+':'),f'my j({expr})']
    return _quote('{')+' & '+' & '.join(pieces)+' & '+_quote('}')


def _native_value(value):
    if value is None: return '""'
    if isinstance(value,bool): return str(value).lower()
    if isinstance(value,(int,float)):
        if not math.isfinite(value) or abs(value)>2**53-1: raise ValueError('Invalid Excel numeric value')
        return str(value)
    if isinstance(value,str):
        if value.lstrip().startswith(('=','+','-','@')):validate_spreadsheet_formula(value)
        if len(value.encode('utf-16-le'))//2>32767: raise ValueError('Excel cell text exceeds limit')
        return _quote(value)
    if isinstance(value,(list,tuple)):
        if sum(len(v) if isinstance(v,(list,tuple)) else 1 for v in value)>10000: raise ValueError('Excel patch is too large')
        return '{'+', '.join(_native_value(v) for v in value)+'}'
    raise ValueError('Unsupported Excel property value')


def _number(value,minimum=0,maximum=100000):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or not minimum<=value<=maximum: raise ValueError('Invalid Excel measurement')
    return str(value)


def _enum(value,table):
    if type(value) is not int or value not in table: raise ValueError('Unsupported Excel enum value')
    return table[value]


def _color(value):
    if type(value) is int and 0<=value<=0xFFFFFF:return '{'+', '.join(str((value >> shift)&255) for shift in (0,8,16))+'}'
    if not isinstance(value,str) or not re.fullmatch('#[0-9a-fA-F]{6}',value): raise ValueError('Color must be #RRGGBB')
    return _rgb(value)


def _normalize(value):
    if isinstance(value,list): return [_normalize(v) for v in value]
    if not isinstance(value,dict): return value
    for key,item in value.items():
        if key in ('color','back_color') and isinstance(item,list) and len(item)==3 and all(isinstance(v,(int,float)) for v in item): value[key]='#'+''.join(f'{max(0,min(255,round(v))):02X}' for v in item)
        elif key in ('horizontal_alignment','vertical_alignment','underline','chart_type') and isinstance(item,str):
            table={'horizontal_alignment':_H_ALIGN,'vertical_alignment':_V_ALIGN,'underline':_UNDERLINE,'chart_type':_CHART}[key]
            value[key]=next((k for k,v in table.items() if v==item),item)
        else: value[key]=_normalize(item)
    return value


class MacExcelSession:
    kind='sheet'
    engine='msoffice'

    @classmethod
    def open_document(cls,path,*,read_only=False,visible=False):
        self=cls.__new__(cls)
        self._source=Path(path).expanduser().resolve()
        self._deadline=time.monotonic()+600
        validate_native_input(self._source,'spreadsheet',deadline=self._deadline)
        root=_container_root('spreadsheet');root.mkdir(parents=True,exist_ok=True,mode=0o700)
        self._lock=OfficeJobLock(root);self._lock.acquire(self._deadline)
        self._job=Path(tempfile.mkdtemp(prefix='session-',dir=root))
        self._native=self._job/('owned-'+uuid4().hex+'.xlsx')
        self._closed=False;self._failed=False;self._attached=False;self._counter=0
        self._read_only=bool(read_only);self._execute=subprocess.run
        self._logical=self._source
        self._logical_state=snapshot_artifact_state(self._source,deadline=self._deadline)
        try:
            copy_file_before_deadline(self._source,self._native,deadline=self._deadline)
            validate_native_input(self._native,'spreadsheet',deadline=self._deadline)
            if snapshot_artifact_state(self._source,deadline=self._deadline) != self._logical_state:
                raise ValueError('Source file changed during Excel staging')
            if snapshot_artifact_state(self._native,deadline=self._deadline).sha256 != self._logical_state.sha256:
                raise ValueError('Source file changed during Excel staging')
            self._run(f'''if {_quote(self._native.name)} is in (name of every workbook) then error "Owned name collision"
set ownedBook to open workbook workbook file name {_quote(str(self._native))} with editable
if full name of ownedBook is not {_quote(str(self._native))} then error "Owned open identity mismatch"
return "{{}}"''',bind=False)
            return self
        except BaseException:
            self._lock.close()
            raise

    @classmethod
    def new_document(cls,*,visible=False):
        self=cls.__new__(cls)
        self._deadline=time.monotonic()+600
        root=_container_root('spreadsheet');root.mkdir(parents=True,exist_ok=True,mode=0o700)
        self._lock=OfficeJobLock(root);self._lock.acquire(self._deadline)
        self._job=Path(tempfile.mkdtemp(prefix='session-',dir=root))
        self._native=self._job/('owned-'+uuid4().hex+'.xlsx');self._source=self._native
        self._logical=None;self._logical_state=None
        self._closed=False;self._failed=False;self._attached=False;self._counter=0
        self._read_only=False;self._execute=subprocess.run;self._sheet_index=1
        try:
            self._run('set ownedBook to make new workbook\nlog "WPSCOMPOSER_CREATED_WORKBOOK:" & (name of ownedBook)\nsave workbook as ownedBook filename '+_quote(str(self._native))+' file format Excel XML file format\nset ownedBook to workbook '+_quote(self._native.name)+'\nif full name of ownedBook is not '+_quote(str(self._native))+' then error "Owned new workbook identity mismatch"\nreturn "{}"',bind=False)
            return self
        except BaseException:
            self._lock.close();raise

    @classmethod
    def attach_active(cls):
        self=cls.__new__(cls)
        self._deadline=time.monotonic()+600
        root=_container_root('spreadsheet');root.mkdir(parents=True,exist_ok=True,mode=0o700)
        self._lock=OfficeJobLock(root);self._lock.acquire(self._deadline)
        self._job=Path(tempfile.mkdtemp(prefix='session-',dir=root))
        self._native=self._job/'unbound.xlsx';self._source=self._native
        self._closed=False;self._failed=False;self._attached=True;self._counter=0;self._read_only=False;self._execute=subprocess.run
        try:
            identity=self._run('if (count of workbooks) is 0 then return "{\\\"active\\\":false}"\nset candidate to active workbook\nreturn '+_object({'active':'true','path':'full name of candidate','name':'name of candidate','read_only':'read only of candidate'}),bind=False)
            if not identity.get('active'):raise ValueError('No active Microsoft Excel workbook')
            native=Path(identity['path'])
            if not native.is_absolute() or not native.is_file():raise NotImplementedError('Unsaved Excel active document identity is not yet safely bindable')
            native=native.resolve()
            validate_native_input(native,'spreadsheet',deadline=self._deadline)
            if native.name!=identity['name']:raise ValueError('Active Excel workbook path and name mismatch')
            self._native=native;self._source=native;self._read_only=identity['read_only']
            self._logical=native
            self._logical_state=snapshot_artifact_state(native,deadline=self._deadline)
            self._run('return "{}"')
            return self
        except BaseException:
            self._closed=True;self._lock.close();raise

    def preflight_save(self,output=None,*,overwrite=False):
        self._assert_live(write=True)
        if self._attached:
            if output is not None:raise NotImplementedError('Attached Excel save-copy cannot preserve native binding')
            if not self._native.is_absolute() or not self._native.is_file() or not os.access(self._native,os.W_OK):raise ValueError('Active Excel workbook requires an existing writable file')
            state=self._run('return '+_object({'read_only':'read only of ownedBook'}))
            if state['read_only']:raise PermissionError('Active Excel workbook is read-only')
        elif output is not None:self._destination(output,'.xlsx')
        return None

    def __enter__(self): return self
    def __exit__(self,*args): self.close(save_changes=False)
    def supports_attached_save_copy(self): return False
    def is_bound_to(self,path): return not self._closed and Path(path).expanduser().resolve()==self._native

    def _assert_live(self,write=False):
        if self._closed or self._failed: raise RuntimeError('Excel session is closed or quarantined')
        if write and self._read_only: raise PermissionError('Excel session is read-only')

    def _run(self,body,*,bind=True):
        self._assert_live()
        remaining=min(60,self._deadline-time.monotonic())
        if remaining<=0:
            self._failed=True
            detail={'component':'spreadsheet','staging_path':str(self._job),'owned_path':str(self._native),'closed':False,'code':'NATIVE_OFFICE_TIMEOUT'}
            (self._job/'recovery.json').write_text(json.dumps(detail)+'\n')
            if self._lock is not None:self._lock.quarantine(detail)
            raise RuntimeError('Excel session deadline expired; owned recovery files retained')
        self._counter+=1
        stem=self._job/f'step-{self._counter:04d}'
        binding=f'''set ownedBook to workbook {_quote(self._native.name)}
if full name of ownedBook is not {_quote(str(self._native))} then error "Owned workbook identity mismatch"''' if bind else ''
        guard=''
        if self._attached:
            guard='if not application '+_quote(_APP)+' is running then '+('error \"Bound Excel application is no longer running\"' if bind else 'return \"{\\\"active\\\":false}\"')+'\n'
        script=_JSON+guard+f'with timeout of {max(1,math.ceil(remaining))} seconds\n tell application {_quote(_APP)}\n'+binding+'\n'+body+'\n end tell\nend timeout\n'
        stem.with_suffix('.applescript').write_text(script)
        stdout=stderr=''
        try:
            result=self._execute(['/usr/bin/osascript',str(stem.with_suffix('.applescript'))],capture_output=True,text=True,timeout=remaining)
            stdout,stderr=result.stdout,result.stderr
            if result.returncode: raise RuntimeError('Native Excel session operation failed')
            return _normalize(json.loads(stdout))
        except BaseException as exc:
            if isinstance(exc,subprocess.TimeoutExpired): stdout,stderr=exc.stdout or '',exc.stderr or ''
            self._failed=True
            detail={'component':'spreadsheet','staging_path':str(self._job),'owned_path':str(self._native),'step':self._counter,'closed':False,'code':'NATIVE_OFFICE_TIMEOUT' if isinstance(exc,subprocess.TimeoutExpired) else 'NATIVE_OFFICE_EXECUTION_FAILED'}
            if getattr(self,'_clipboard_operation',False):detail.update(clipboard_changed=True,clipboard_may_have_changed=True)
            (self._job/'recovery.json').write_text(json.dumps(detail)+'\n')
            if self._lock is not None: self._lock.quarantine(detail)
            error=RuntimeError('Native Excel session failed; owned recovery files retained')
            error.staging_path=str(self._job);error.diagnostic_path=str(self._job/'recovery.json')
            if getattr(self,'_clipboard_operation',False):error.clipboard_changed=True;error.clipboard_may_have_changed=True
            raise error from None
        finally:
            for name,text in [('stdout',stdout),('stderr',stderr)]:
                stem.with_suffix('.'+name+'.log').write_text(text.decode(errors='replace') if isinstance(text,bytes) else text)

    def _range_json(self,range_expr,id_expr):
        props={'id':id_expr,'address':'get address rng','value':'value of rp','formula':'formula of rp','display_text':'string value of rp','number_format':'number format of rp','horizontal_alignment':'horizontal alignment of rp','vertical_alignment':'vertical alignment of rp','wrap_text':'wrap text of rp','indent':'indent level of rp','row_height':'row height of rp','column_width':'column width of rp','merged':'merge cells of rp'}
        font={'name':'name of fp','size':'font size of fp','bold':'bold of fp','italic':'italic of fp','underline':'underline of fp','strikethrough':'strikethrough of fp','color':'color of fp'}
        return f'''set rng to {range_expr}
set rp to properties of rng
set fp to properties of font object of rng
set rangeJSON to {_object(props)}
set fontJSON to {_object(font)}
set fillJSON to {_object({'color':'color of interior object of rng'})}
set mergeJSON to "null"
if merge cells of rp is true then set mergeJSON to my j(get address merge area of rng)
set rangeJSON to (text 1 thru -2 of rangeJSON) & ",\\"font\\":" & fontJSON & ",\\"fill\\":" & fillJSON & ",\\"merge_area\\":" & mergeJSON & "}}"
'''

    def inspect_document(self,include_cells=True,max_cells=1000,include_empty=False):
        if type(max_cells) is not int or not 0<=max_cells<=100000 or type(include_cells) is not bool or type(include_empty) is not bool: raise ValueError('Invalid Excel inspection limit')
        cell_code=self._range_json('cell (firstRow + r - 1) of column (firstCol + c - 1) of ws','"sheet:" & sheetIndex & "/cell:" & (get address rng)')
        sheet_fields={'id':'"sheet:" & sheetIndex','index':'sheetIndex','name':'name of ws','visible':'(visible of ws) as text','used_range':'get address usedRng','used_rows':'usedRows','used_columns':'usedCols'}
        page_fields={'top_margin':'top margin of pp','bottom_margin':'bottom margin of pp','left_margin':'left margin of pp','right_margin':'right margin of pp','orientation':'(page orientation of pp) as text','zoom':'zoom of pp','fit_to_pages_wide':'fit to pages wide of pp','fit_to_pages_tall':'fit to pages tall of pp','print_area':'print area of pp','left_header':'left header of pp','center_header':'center header of pp','right_header':'right header of pp'}
        geom={'left':'left position of cp','top':'top of cp','width':'width of cp','height':'height of cp'}
        chart_fields={'id':'"sheet:" & sheetIndex & "/chart:" & chartIndex','index':'chartIndex','name':'name of cp','chart_type':'(chart type of chart of co) as text','has_title':'has title of chart of co','title':'titleText'}
        shape_fields={'id':'"sheet:" & sheetIndex & "/shape:@name=" & (name of sp)','index':'shapeIndex','shape_id':'missing value','name':'name of sp','type':'(shape type of sp) as text'}
        shape_fill={'visible':'visible of sf','type':'(fill format type of sf) as text','color':'fore color of sf','back_color':'back color of sf','transparency':'transparency of sf'}
        shape_line={'visible':'visible of sl','color':'fore color of sl','weight':'weight of sl','dash_style':'(dash style of sl) as text','transparency':'transparency of sl'}
        shape_geom={'left':'left position of sp','top':'top of sp','width':'width of sp','height':'height of sp','rotation':'rotation of sp','z_order':'z order position of sp'}
        root_fields={'kind':'"sheet"','name':'name of ownedBook','path':_quote(str(self._source)),'saved':'saved of ownedBook','sheet_count':'count of worksheets of ownedBook','cells_truncated':'truncated'}
        body=f'''set sheetJSONs to {{}}
set remainingCells to {max_cells}
set examined to 0
set truncated to false
repeat with sheetIndex from 1 to (count of worksheets of ownedBook)
 set ws to worksheet sheetIndex of ownedBook
 activate object ws
 set usedRng to used range of ws
 set usedRows to count of rows of usedRng
 set usedCols to count of columns of usedRng
 set firstRow to first row index of usedRng
 set firstCol to first column index of usedRng
 set cellJSONs to {{}}
 if {str(include_cells).lower()} then
  repeat with r from 1 to usedRows
   repeat with c from 1 to usedCols
    if remainingCells <= 0 or examined >= 100000 then
     set truncated to true
     exit repeat
    end if
    set examined to examined + 1
    set cellValue to value of cell (firstRow + r - 1) of column (firstCol + c - 1) of ws
    if {str(include_empty).lower()} or cellValue is not "" then
     {cell_code}
     set end of cellJSONs to rangeJSON
     set remainingCells to remainingCells - 1
    end if
   end repeat
   if truncated then exit repeat
  end repeat
 end if
 set pp to properties of page setup object of ws
 set pageJSON to {_object(page_fields)}
 set chartJSONs to {{}}
 repeat with chartIndex from 1 to (count of chart objects of ws)
  set co to chart object chartIndex of ws
  set cp to properties of co
  set titleText to missing value
  if has title of chart of co then set titleText to chart title text of chart title of chart of co
  set chartJSON to {_object(chart_fields)}
  set geometryJSON to {_object(geom)}
  set end of chartJSONs to (text 1 thru -2 of chartJSON) & ",\\"geometry\\":" & geometryJSON & "}}"
 end repeat
 set shapeJSONs to {{}}
 repeat with shapeIndex from 1 to (count of shapes of ws)
  set shp to shape shapeIndex of ws
  set sp to properties of shp
  set sf to properties of fill format of shp
  set sl to properties of line format of shp
  set fillJSON to {_object(shape_fill)}
  set lineJSON to {_object(shape_line)}
  set shapeJSON to {_object(shape_fields)}
  set geometryJSON to {_object(shape_geom)}
  set end of shapeJSONs to (text 1 thru -2 of shapeJSON) & ",\\"geometry\\":" & geometryJSON & ",\\"fill\\":" & fillJSON & ",\\"line\\":" & lineJSON & "}}"
 end repeat
 set freezeJSON to my j(freeze panes of window 1 of ownedBook)
 set sheetJSON to {_object(sheet_fields)}
 set end of sheetJSONs to (text 1 thru -2 of sheetJSON) & ",\\"page_setup\\":" & pageJSON & ",\\"freeze_panes\\":" & freezeJSON & ",\\"cells\\":[" & my joined(cellJSONs) & "],\\"shapes\\":[" & my joined(shapeJSONs) & "],\\"charts\\":[" & my joined(chartJSONs) & "]}}"
end repeat
set resultJSON to {_object(root_fields)}
return (text 1 thru -2 of resultJSON) & ",\\"sheets\\":[" & my joined(sheetJSONs) & "]}}"'''
        if self._attached:
            body='if full name of active workbook is not '+_quote(str(self._native))+' then error \"Active workbook changed before inspection\"\nset originalSheetName to name of active sheet of ownedBook\nset originalSelectionAddress to get address selection\ntry\n'+body
            restore='activate object worksheet originalSheetName of ownedBook\nselect range originalSelectionAddress of worksheet originalSheetName of ownedBook\n'
            body=body.replace('return (text 1 thru -2 of resultJSON)',restore+'return (text 1 thru -2 of resultJSON)')
            body+='\non error inspectionError number inspectionNumber\nif inspectionNumber is not -1712 then\nif full name of ownedBook is '+_quote(str(self._native))+' then\n'+restore+'end if\nend if\nerror inspectionError number inspectionNumber\nend try'
        return self._run(body)

    def inspect_selection(self):
        return self._run(f'''if full name of active workbook is not {_quote(str(self._native))} then error "Active selection belongs to another workbook"
{self._range_json('selection','"selection"')}
return rangeJSON''')

    def _target_script(self,target):
        sheet,kind,ref=parse_target(target)
        if kind=='selection': return f'if full name of active workbook is not {_quote(str(self._native))} then error "Selection owner mismatch"\nset obj to selection',kind
        prefix=f'set ws to worksheet {sheet} of ownedBook\nactivate object ws\n'
        expressions={'sheet':'ws','cell':f'range {_quote(str(ref))} of ws','range':f'range {_quote(str(ref))} of ws','shape':f'shape {ref} of ws','shape_name':f'shape {_quote(str(ref))} of ws','chart':f'chart object {ref} of ws'}
        if kind=='shape_name':
            prefix += 'set shapeMatches to 0\nrepeat with shapeIndex from 1 to (count of shapes of ws)\n if name of shape shapeIndex of ws is '+_quote(str(ref))+' then set shapeMatches to shapeMatches + 1\nend repeat\nif shapeMatches is not 1 then error \"Shape name is missing or ambiguous\"\n'
        return prefix+'set obj to '+expressions[kind],kind

    def apply_format_patch(self,target,**patch):
        self._assert_live(write=True)
        prefix,kind=self._target_script(target)
        from .edit_preflight import compile_excel_patch
        operations, range_kind = compile_excel_patch(target, patch)
        lines=[prefix,'set acceptedKeys to {}','set rejectedKeys to {}']
        # Native failure aborts the session; do not hide timeout/late-writer errors
        # in per-property try blocks. Validation already happened for every key.
        for key,command in operations: lines += [command,f'set end of acceptedKeys to {_quote(key)}']
        if range_kind: lines += ['calculate obj']
        lines += ['return '+_object({'accepted':'acceptedKeys','rejected':'rejectedKeys'})]
        self._fresh_empty_sheets=set()
        return self._run('\n'.join(lines))

    def apply_structural_op(self,op):
        self._assert_live(write=True)
        if not isinstance(op,dict): raise ValueError('Invalid Excel structural operation')
        from .edit_preflight import validate_excel_structural
        validate_excel_structural(op)
        verb=op.get('op');target=op.get('target');etype=op.get('type')
        if verb in ('move','clone'):
            self._fresh_empty_sheets=set()
            sheet,kind,ref=parse_target(target)
            prefix,_=self._target_script(target)
            to=op.get('to')
            if kind=='sheet':
                if to is None:
                    side='after';destination='(count of worksheets of ownedBook)' if verb=='move' else str(sheet)
                elif to=='end':
                    side='after';destination='(count of worksheets of ownedBook)'
                elif isinstance(to,dict) and len(to)==1 and next(iter(to)) in ('before','after'):
                    side=next(iter(to));index=to[side]
                    if type(index) is not int or index<1:raise ValueError('Invalid worksheet destination')
                    destination=str(index)
                else:raise ValueError('Invalid worksheet destination')
                lines=[prefix,'set sourceSheetName to name of obj','set beforeSheetNames to name of every worksheet of ownedBook']
                if verb=='clone' and to=='end':
                    # Capture the pre-copy destination. Evaluating ``count``
                    # after insertion points one worksheet past the clone.
                    lines += ['set destinationSheetIndex to count of worksheets of ownedBook']
                    destination='destinationSheetIndex'
                if verb=='move':
                    lines += [f'if name of worksheet {destination} of ownedBook is not sourceSheetName then',f' move obj to {side} worksheet {destination} of ownedBook','end if','set resultingSheetName to sourceSheetName']
                else:
                    resulting_index=f'({destination} + 1)' if side=='after' else destination
                    lines += ['set beforeSheetCount to count of worksheets of ownedBook',f'copy worksheet obj {side} worksheet {destination} of ownedBook','if (count of worksheets of ownedBook) is not beforeSheetCount + 1 then error "Worksheet clone count mismatch"',f'set resultingSheetName to name of worksheet {resulting_index} of ownedBook']
                lines += ['set resultingIndex to entry_index of worksheet resultingSheetName of ownedBook','return '+_object({'type':'"sheet"','moved' if verb=='move' else 'cloned':'true','from':_quote(target),'path':'"sheet:" & resultingIndex'})]
                result=self._run('\n'.join(lines))
                destination_index=int(result['path'].split(':')[1]);selected=getattr(self,'_sheet_index',1)
                if verb=='clone':
                    if selected>=destination_index:selected+=1
                elif selected==sheet:selected=destination_index
                elif sheet<selected<=destination_index:selected-=1
                elif destination_index<=selected<sheet:selected+=1
                self._sheet_index=selected
                return result
            if kind not in ('cell','range'):raise ValueError('Unsupported Excel move/clone target')
            axis=op.get('axis','row')
            if axis not in ('row','column') or not isinstance(to,dict) or set(to)!={'index'} or type(to['index']) is not int:raise ValueError('Invalid row or column destination')
            dest=to['index'];match=re.match(r'([A-Z]+)(\d+)',ref)
            origin=int(match[2]) if axis=='row' else _column_number(match[1])
            if not 1<=dest<=(1048576 if axis=='row' else 16384):raise ValueError('Destination exceeds Excel bounds')
            if verb=='move' and dest in (origin,origin+1):return {'type':axis,'moved':False,'from':target,axis:origin,'clipboard_changed':False}
            destination=f'{dest}:{dest}' if axis=='row' else f'{_column_name(dest)}:{_column_name(dest)}'
            result_index=dest-1 if verb=='move' and dest>origin else dest
            prefix+='\nset obj to range '+_quote(ref.split(':')[0])+' of ws'
            lines=[prefix,('cut' if verb=='move' else 'copy')+' range (entire '+axis+' of obj)',f'insert into range (range {_quote(destination)} of ws)','set cut copy mode to false','return '+_object({'type':_quote(axis),'moved' if verb=='move' else 'cloned':'true','from':_quote(target),axis:str(result_index),'clipboard_changed':'true','requested_index':str(dest)})]
            self._clipboard_operation=True
            try:return self._run('\n'.join(lines))
            finally:self._clipboard_operation=False
        if verb=='insert' and etype=='sheet':
            props=op.get('props') or {}
            if set(props)-{'name'}: raise ValueError('Unsupported worksheet insert fields')
            name=props.get('name','Sheet-'+uuid4().hex[:8]);quoted=_sheet_name(name)
            body=f'''if {quoted} is in (name of every worksheet of ownedBook) then error "Duplicate sheet name"
tell ownedBook
 make new worksheet at end with properties {{name:{quoted}}}
end tell
return {_object({'type':'"sheet"','path':'"sheet:" & (count of worksheets of ownedBook)'})}'''
            result=self._run(body)
            self._fresh_empty_sheets=getattr(self,'_fresh_empty_sheets',set())|{int(result['path'].split(':')[1])}
            return result
        if verb=='insert':
            self._fresh_empty_sheets=set()
            if etype not in ('row','column'): raise ValueError('Unsupported Excel insert type')
            sheet,kind,ref=parse_target(op.get('parent'))
            if kind not in ('sheet','cell','range'): raise ValueError('Invalid Excel structural parent')
            position=op.get('position','end');props=op.get('props') or {}
            if set(props)-{'values'}: raise ValueError('Unsupported Excel insert fields')
            if isinstance(position,dict):
                if set(position)!={'index'} or type(position['index']) is not int: raise ValueError('Invalid insert position')
                index=position['index']
            elif ref:
                match=re.match(r'([A-Z]+)(\d+)',ref);index=int(match[2]) if etype=='row' else _column_number(match[1])
            elif position=='end':
                snap=self.inspect_document(include_cells=False);entry=snap['sheets'][sheet-1]
                match=re.match(r'\$?([A-Z]+)\$?(\d+)',entry['used_range']);index=(int(match[2])+entry['used_rows']) if etype=='row' else (_column_number(match[1])+entry['used_columns'])
            else: raise ValueError('Invalid insert position')
            if not 1<=index<=(1048576 if etype=='row' else 16384): raise ValueError('Insert index outside Excel bounds')
            address=f'{index}:{index}' if etype=='row' else f'{_column_name(index)}:{_column_name(index)}'
            lines=[f'set ws to worksheet {sheet} of ownedBook','activate object ws',f'insert into range (range {_quote(address)} of ws)']
            values=props.get('values')
            if values:
                if not isinstance(values,(list,tuple)) or len(values)>10000: raise ValueError('Invalid inserted values')
                for offset,value in enumerate(values,1):
                    addr=f'{_column_name(offset)}{index}' if etype=='row' else f'{_column_name(index)}{offset}'
                    lines += [f'set value of range {_quote(addr)} of ws to {_native_value(value)}']
            lines += ['return '+_object({'type':_quote(etype),'path':_quote(f'sheet:{sheet}'),etype:str(index)})]
            return self._run('\n'.join(lines))
        if verb=='remove':
            sheet,kind,ref=parse_target(target)
            prefix,_=self._target_script(target)
            if kind=='sheet':
                if parse_target(target)[0] not in getattr(self,'_fresh_empty_sheets',set()):raise NotImplementedError('Deleting an existing Excel worksheet requires an unverified confirmation-free native primitive')
                self._fresh_empty_sheets=set()
                command='if (count of worksheets of ownedBook) <= 1 then error "Cannot remove last worksheet"\ndelete obj'
            elif kind in ('cell','range'):
                axis=op.get('axis','row')
                if axis not in ('row','column'): raise ValueError('Invalid Excel removal axis')
                prefix+='\nset obj to range '+_quote(parse_target(target)[2].split(':')[0])+' of ws'
                command='delete range (entire '+axis+' of obj)'
            elif kind in ('shape','shape_name','chart'): command='delete obj'
            else: raise ValueError('Unsupported Excel removal target')
            return self._run(prefix+'\n'+command+'\nreturn '+_object({'removed':_quote(target)}))
        raise ValueError('Unsupported Excel structural operation')

    def _business_sheet(self):
        self._fresh_empty_sheets=set()
        index=getattr(self,'_sheet_index',1)
        return f'sheet:{index}'

    def select_sheet(self,index):
        if type(index) is not int or index<1:raise ValueError('Invalid worksheet index')
        body,_=self._target_script(f'sheet:{index}')
        result=self._run(body+'\nreturn '+_object({'path':_quote(f'sheet:{index}'),'name':'name of ws'}))
        self._sheet_index=index
        return result

    def rename_sheet(self,index,name):
        return self.apply_format_patch(f'sheet:{index}',name=name)

    def add_sheet(self,name=None):
        result=self.apply_structural_op({'op':'insert','type':'sheet','props':{'name':name} if name else {}})
        self._sheet_index=int(result['path'].split(':')[1])
        return result

    def _cell_address(self,row,col):
        if type(row) is not int or type(col) is not int or not 1<=row<=1048576 or not 1<=col<=16384:raise ValueError('Invalid Excel cell coordinates')
        return _column_name(col)+str(row)

    def write_cell(self,row,col,value):
        return self.apply_format_patch(self._business_sheet()+'/cell:'+self._cell_address(row,col),value='' if value is None else value)

    def set_formula(self,row,col,formula):
        return self.apply_format_patch(self._business_sheet()+'/cell:'+self._cell_address(row,col),formula=formula)

    def write_table(self,start_row,start_col,data,header_bold=True,header_shade='#4472C4',header_font_color='#FFFFFF',font_size=11):
        self._assert_live(write=True)
        if not data:return
        if not isinstance(data,(list,tuple)) or len(data)>10000:raise ValueError('Invalid Excel table')
        commands=[]
        for r,row in enumerate(data):
            if not isinstance(row,(list,tuple)):raise ValueError('Invalid Excel table row')
            for c,value in enumerate(row):
                address=self._cell_address(start_row+r,start_col+c)
                commands += [f'set obj to range {_quote(address)} of ws',f'set value of obj to {_native_value(value)}',f'set font size of font object of obj to {_number(font_size,1,409)}']
                if r==0:
                    if header_bold:commands += ['set bold of font object of obj to true']
                    if header_shade:commands += ['set color of interior object of obj to '+_color(header_shade)]
                    if header_font_color:commands += ['set color of font object of obj to '+_color(header_font_color)]
        if len(commands)>50000:raise ValueError('Excel table exceeds native operation limit')
        prefix,_=self._target_script(self._business_sheet())
        return self._run(prefix+'\n'+'\n'.join(commands)+'\ncalculate ws\nreturn "{}"')

    def set_cell_style(self,row,col,bold=None,italic=None,font_size=None,font_color=None,fill_color=None,align=None,number_format=None):
        return self.set_range_style(self._cell_address(row,col),bold,italic,font_size,font_color,fill_color,align,number_format)

    def set_range_style(self,range_str,bold=None,italic=None,font_size=None,font_color=None,fill_color=None,align=None,number_format=None):
        font={k:v for k,v in {'bold':bold,'italic':italic,'size':font_size,'color':font_color}.items() if v is not None}
        patch={}
        if font:patch['font']=font
        if fill_color is not None:patch['fill']={'color':fill_color}
        if align is not None:patch['horizontal_alignment']=align
        if number_format is not None:patch['number_format']=number_format
        return self.apply_format_patch(self._business_sheet()+'/range:'+range_str,**patch)

    def set_borders(self,range_str,style=1,weight=2,color='#000000'):
        return self.apply_format_patch(self._business_sheet()+'/range:'+range_str,borders={str(edge):{'style':style,'weight':weight,'color':color} for edge in (7,8,9,10,11,12)})

    def merge_cells(self,range_str):
        self._assert_live(write=True)
        prefix,_=self._target_script(self._business_sheet()+'/range:'+range_str)
        return self._run(prefix+'\nmerge obj\nreturn "{}"')

    def freeze_panes(self,cell):
        self._assert_live(write=True)
        prefix,_=self._target_script(self._business_sheet()+'/cell:'+cell)
        return self._run(prefix+'\nset ownedWindow to window 1 of ownedBook\nset freeze panes of ownedWindow to false\nset split column of ownedWindow to (first column index of obj) - 1\nset split row of ownedWindow to (first row index of obj) - 1\nselect obj\nset freeze panes of ownedWindow to true\nreturn '+_object({'freeze_panes':'freeze panes of ownedWindow','split_column':'split column of ownedWindow','split_row':'split row of ownedWindow'}))

    def set_column_width(self,col,width):
        return self.apply_format_patch(self._business_sheet()+'/cell:'+self._cell_address(1,col),column_width=width)

    def set_row_height(self,row,height):
        return self.apply_format_patch(self._business_sheet()+'/cell:'+self._cell_address(row,1),row_height=height)

    def autofit(self):
        self._assert_live(write=True)
        prefix,_=self._target_script(self._business_sheet())
        return self._run(prefix+'\nautofit entire column of used range of ws\nreturn "{}"')

    def add_chart(self,chart_type=4,left=100,top=100,width=400,height=300,source_range=None,title=None):
        self._assert_live(write=True)
        chart_type=_enum(chart_type,_CHART)
        geometry={'left position':left,'top':top,'width':width,'height':height}
        props=', '.join(k+':'+_number(v,0 if k in ('left position','top') else 1) for k,v in geometry.items())
        source=''
        if source_range is not None:
            _,_,address=parse_target(self._business_sheet()+'/range:'+source_range)
            source=f'set source data nativeChart source range {_quote(address)} of ws plot by columns\n'
        title_source='' if title is None else 'set has title of chart of co to true\nset chart title text of chart title of chart of co to '+_quote(title)+'\n'
        prefix,_=self._target_script(self._business_sheet())
        body=prefix+'\nset chartIndex to (count of chart objects of ws) + 1\ntell ws\nmake new chart object at end with properties {'+props+'}\nend tell\nset co to chart object chartIndex of ws\nset nativeChart to chart of co\n'+source+'set chart type of nativeChart to '+chart_type+'\n'+title_source+'return '+_object({'path':_quote(self._business_sheet()+'/chart:')+' & chartIndex','name':'name of co'})
        return self._run(body)

    def add_title_row(self,range_str,text,fill_color='#4472C4',font_color='#FFFFFF',size=14):
        self._assert_live(write=True)
        prefix,_=self._target_script(self._business_sheet()+'/range:'+range_str)
        value=_native_value(text);shade=_color(fill_color);color=_color(font_color);size=_number(size,1,409)
        return self._run(prefix+'\nmerge obj\nset value of obj to '+value+'\nset bold of font object of obj to true\nset font size of font object of obj to '+size+'\nset color of font object of obj to '+color+'\nset color of interior object of obj to '+shade+'\nset horizontal alignment of obj to horizontal align center\nreturn "{}"')

    def conditional_format(self,range_str,rule_type='cellvalue',operator=5,formula='0',fill_color='#FFC7CE'):
        self._assert_live(write=True)
        prefix,_=self._target_script(self._business_sheet()+'/range:'+range_str)
        rule=_enum({'cellvalue':1,'expression':2}.get(rule_type,rule_type),{1:'cell value',2:'expression'})
        operator=_enum(operator,{1:'operator between',2:'operator not between',3:'operator equal',4:'operator not equal',5:'operator greater',6:'operator less',7:'operator greater equal',8:'operator less equal'})
        validate_spreadsheet_formula(formula);formula=_quote(formula);color=_color(fill_color)
        return self._run(prefix+'\nrepeat (count of format conditions of obj) times\ndelete format condition 1 of obj\nend repeat\nmake new format condition at obj with properties {format condition type:'+rule+', condition operator:'+operator+', formula 1:'+formula+'}\nset color of interior object of format condition 1 of obj to '+color+'\nreturn '+_object({'condition_count':'count of format conditions of obj'}))

    def set_header_footer(self,left=None,center=None,right=None):
        self._assert_live(write=True)
        commands=[]
        for key,value in [('left',left),('center',center),('right',right)]:
            if value:commands.append('set '+key+' header of page setup object of ws to '+_quote(value))
        prefix,_=self._target_script(self._business_sheet())
        return self._run(prefix+'\n'+'\n'.join(commands)+'\nreturn "{}"')

    def save_xlsx(self,path):return self.save(path)

    def _destination(self,path,suffix):
        destination=Path(path).expanduser().resolve()
        if destination==self._source: raise ValueError('File sessions never overwrite their source')
        if destination.suffix.lower()!=suffix: raise ValueError('Unexpected Excel output extension')
        if destination.exists(): raise FileExistsError(destination)
        return destination

    def save_current(self):
        self._assert_live(write=True)
        if self._attached:self.preflight_save()
        if not self._attached and self._logical is None:
            raise ValueError('New Excel session requires an explicit save destination')
        self._run('save ownedBook\nreturn "{}"')
        if self._attached:
            self._logical_state=snapshot_artifact_state(self._logical,deadline=self._deadline)
            return str(self._logical)
        validate_before_deadline(ValidatorSpec.from_callable(validate_office_package,'xlsx'),self._native,self._deadline)
        publish_artifact(
            self._native,self._logical,overwrite=True,
            validator=lambda path: validate_before_deadline(ValidatorSpec.from_callable(validate_office_package,'xlsx'),path,self._deadline),
            deadline=self._deadline,expected_destination=self._logical_state,
        )
        self._logical_state=snapshot_artifact_state(self._logical,deadline=self._deadline)
        return str(self._logical)

    def save(self,destination):
        self._assert_live()
        if self._attached:raise NotImplementedError('Attached Excel save-copy cannot preserve native binding')
        destination=self._destination(destination,'.xlsx')
        self._run('save ownedBook\nreturn "{}"')
        validate_before_deadline(ValidatorSpec.from_callable(validate_office_package,'xlsx'),self._native,self._deadline)
        destination.parent.mkdir(parents=True,exist_ok=True)
        publish_artifact(self._native,destination,overwrite=False,validator=lambda path: validate_before_deadline(ValidatorSpec.from_callable(validate_office_package,'xlsx'),path,self._deadline),deadline=self._deadline)
        self._logical=destination
        self._logical_state=snapshot_artifact_state(destination,deadline=self._deadline)
        return str(destination)

    def save_copy(self,destination):
        logical, logical_state = self._logical, self._logical_state
        try:
            return self.save(destination)
        finally:
            self._logical, self._logical_state = logical, logical_state

    def export_pdf(self,destination):
        self._assert_live()
        if self._attached:raise NotImplementedError('Attached Excel export has no verified non-rebinding primitive')
        destination=self._destination(destination,'.pdf')
        native_pdf=self._job/('export-'+uuid4().hex+'.pdf')
        self._run(f'save workbook as ownedBook filename {_quote(str(native_pdf))} file format PDF file format\nif full name of ownedBook is not {_quote(str(self._native))} then error "PDF export changed binding"\nreturn "{{}}"')
        validate_before_deadline(ValidatorSpec.from_callable(validate_pdf),native_pdf,self._deadline)
        destination.parent.mkdir(parents=True,exist_ok=True)
        publish_artifact(native_pdf,destination,overwrite=False,validator=lambda path: validate_before_deadline(ValidatorSpec.from_callable(validate_pdf),path,self._deadline),deadline=self._deadline)
        return str(destination)

    def close(self,save_changes=False):
        if self._closed: return
        try:
            if not self._failed and not self._attached:
                if save_changes: self.save_current()
                self._run('close ownedBook saving no\nreturn "{}"')
        finally:
            self._closed=True
            if self._lock is not None: self._lock.close()
            # Keep scripts/evidence and the only edited recovery file; parent may
            # remove this private job explicitly after publication acceptance.
