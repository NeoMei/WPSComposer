"""Owned macOS Excel file sessions with native inspection and bounded mutation.

Files are copied into Excel's container. A held per-app lock spans the session;
all AppleEvents bind the exact owned path. Failed/uncertain steps quarantine the
app and retain staging. Active attachment and clipboard-dependent row/column
move/clone are explicit capability gaps, never simulated with cell rewriting.
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

from ..artifact_transport import copy_file_before_deadline, validate_office_package, validate_pdf, validate_before_deadline, ValidatorSpec
from .input_validation import validate_native_input
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
    if not isinstance(value,str) or not re.fullmatch('#[0-9a-fA-F]{6}',value): raise ValueError('Color must be #RRGGBB')
    return _rgb(value)


def _normalize(value):
    if isinstance(value,list): return [_normalize(v) for v in value]
    if not isinstance(value,dict): return value
    for key,item in value.items():
        if key=='color' and isinstance(item,list) and len(item)==3 and all(isinstance(v,(int,float)) for v in item): value[key]='#'+''.join(f'{max(0,min(255,round(v))):02X}' for v in item)
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
        try:
            copy_file_before_deadline(self._source,self._native,deadline=self._deadline)
            validate_native_input(self._native,'spreadsheet',deadline=self._deadline)
            self._run(f'''if {_quote(self._native.name)} is in (name of every workbook) then error "Owned name collision"
set ownedBook to open workbook workbook file name {_quote(str(self._native))} with editable
if full name of ownedBook is not {_quote(str(self._native))} then error "Owned open identity mismatch"
return "{{}}"''',bind=False)
            return self
        except BaseException:
            self._lock.close()
            raise

    @classmethod
    def attach_active(cls):
        raise NotImplementedError('Mac Excel active document identity and non-rebinding save-copy are not yet verified')

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
        script=_JSON+f'with timeout of {max(1,math.ceil(remaining))} seconds\n tell application {_quote(_APP)}\n'+binding+'\n'+body+'\n end tell\nend timeout\n'
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
            (self._job/'recovery.json').write_text(json.dumps(detail)+'\n')
            if self._lock is not None: self._lock.quarantine(detail)
            error=RuntimeError('Native Excel session failed; owned recovery files retained')
            error.staging_path=str(self._job);error.diagnostic_path=str(self._job/'recovery.json')
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
        page_fields={'top_margin':'top margin of pp','bottom_margin':'bottom margin of pp','left_margin':'left margin of pp','right_margin':'right margin of pp','orientation':'(page orientation of pp) as text','zoom':'zoom of pp','fit_to_pages_wide':'fit to pages wide of pp','fit_to_pages_tall':'fit to pages tall of pp','print_area':'print area of pp'}
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
  set end of shapeJSONs to (text 1 thru -2 of shapeJSON) & ",\\"geometry\\":" & geometryJSON & "}}"
 end repeat
 set sheetJSON to {_object(sheet_fields)}
 set end of sheetJSONs to (text 1 thru -2 of sheetJSON) & ",\\"page_setup\\":" & pageJSON & ",\\"freeze_panes\\":null,\\"cells\\":[" & my joined(cellJSONs) & "],\\"shapes\\":[" & my joined(shapeJSONs) & "],\\"charts\\":[" & my joined(chartJSONs) & "]}}"
end repeat
set resultJSON to {_object(root_fields)}
return (text 1 thru -2 of resultJSON) & ",\\"sheets\\":[" & my joined(sheetJSONs) & "]}}"'''
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
        return prefix+'set obj to '+expressions[kind],kind

    def apply_format_patch(self,target,**patch):
        self._assert_live(write=True)
        prefix,kind=self._target_script(target)
        operations=[]
        def add(key,prop,value): operations.append((key,f'set {prop} to {value}'))
        range_kind=kind in ('cell','range','selection')
        allowed={'value','formula','font','fill','number_format','horizontal_alignment','vertical_alignment','wrap_text','indent','row_height','column_width','borders'} if range_kind else ({'name','page_setup'} if kind=='sheet' else {'geometry','name','chart_type','chart_title'} if kind=='chart' else {'geometry','name','fill','line'})
        if set(patch)-allowed: raise ValueError('Unsupported Excel patch fields')
        for key,value in patch.items():
            if value is None: continue
            if key in ('value','formula'): add(key,key+' of obj',_native_value(value))
            elif key=='font':
                mapping={'name':'name','size':'font size','bold':'bold','italic':'italic','underline':'underline','strikethrough':'strikethrough','color':'color'}
                if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel font fields')
                for sub,item in value.items():
                    val=_color(item) if sub=='color' else _number(item,1,409) if sub=='size' else _enum(item,_UNDERLINE) if sub=='underline' else _quote(item) if sub=='name' else _native_value(item) if type(item) is bool else None
                    if val is None: raise ValueError('Invalid Excel font value')
                    add('font.'+sub,mapping[sub]+' of font object of obj',val)
            elif key in ('fill','line'):
                if range_kind:
                    if key!='fill' or not isinstance(value,dict) or set(value)-{'color'}: raise ValueError('Unsupported Excel range fill fields')
                    if 'color' in value:add('fill.color','color of interior object of obj',_color(value['color']))
                else:
                    mapping={'color':'fore color','back_color':'back color','transparency':'transparency','visible':'visible'}
                    if key=='line':mapping['weight']='weight'
                    if not isinstance(value,dict) or set(value)-set(mapping):raise ValueError('Unsupported Excel shape format fields')
                    for sub,item in value.items():
                        if sub in ('color','back_color'):val=_color(item)
                        elif sub=='visible':
                            if type(item) is not bool:raise ValueError('Shape visibility must be boolean')
                            val=_native_value(item)
                        else:val=_number(item,0,1 if sub=='transparency' else 1000)
                        add(key+'.'+sub,mapping[sub]+' of '+key+' format of obj',val)
            elif key=='geometry':
                mapping={'left':'left position','top':'top','width':'width','height':'height','rotation':'rotation'}
                if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel geometry fields')
                for sub,item in value.items(): add('geometry.'+sub,mapping[sub]+' of obj',_number(item))
            elif key=='name': add(key,'name of obj',_sheet_name(value) if kind=='sheet' else _quote(value))
            elif key=='chart_type': add(key,'chart type of chart of obj',_enum(value,_CHART))
            elif key=='chart_title':
                operations.append((key,'set has title of chart of obj to true\nset chart title text of chart title of chart of obj to '+_quote(value)))
            elif key=='borders':
                if not isinstance(value,dict): raise ValueError('Invalid Excel borders')
                for edge,spec in value.items():
                    try: native_edge=_enum(int(edge),_BORDER)
                    except (TypeError,ValueError): raise ValueError('Invalid border edge') from None
                    if not isinstance(spec,dict) or set(spec)-{'style','weight','color'}: raise ValueError('Unsupported border properties')
                    commands=[f'set ownedBorder to get border obj which border {native_edge}']
                    for sub,item in spec.items(): commands += [f'set '+{'style':'line style','weight':'weight','color':'color'}[sub]+' of ownedBorder to '+(_color(item) if sub=='color' else _enum(item,_BORDER_STYLE if sub=='style' else _BORDER_WEIGHT))]
                    operations.append(('borders.'+str(edge),'\n'.join(commands)))
            elif key=='page_setup':
                mapping={'orientation':'page orientation','top_margin':'top margin','bottom_margin':'bottom margin','left_margin':'left margin','right_margin':'right margin','header_margin':'header margin','footer_margin':'footer margin','zoom':'zoom','fit_to_pages_wide':'fit to pages wide','fit_to_pages_tall':'fit to pages tall','print_area':'print area','print_title_rows':'print title rows','print_title_columns':'print title columns'}
                if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel page setup fields')
                for sub,item in value.items():
                    val=_enum(item,{1:'portrait',2:'landscape'}) if sub=='orientation' else _quote(item) if sub.startswith('print_') else 'false' if sub=='zoom' and item is False else _number(item)
                    add('page_setup.'+sub,mapping[sub]+' of page setup object of obj',val)
            else:
                prop={'number_format':'number format','horizontal_alignment':'horizontal alignment','vertical_alignment':'vertical alignment','wrap_text':'wrap text','indent':'indent level','row_height':'row height','column_width':'column width'}[key]
                if key=='number_format': val=_quote(value)
                elif key=='horizontal_alignment': val=_enum(value,_H_ALIGN)
                elif key=='vertical_alignment': val=_enum(value,_V_ALIGN)
                elif key=='wrap_text':
                    if type(value) is not bool: raise ValueError('wrap_text must be boolean')
                    val=_native_value(value)
                else: val=_number(value,0,255 if key=='column_width' else 15 if key=='indent' else 409)
                add(key,prop+' of obj',val)
        lines=[prefix,'set acceptedKeys to {}','set rejectedKeys to {}']
        # Native failure aborts the session; do not hide timeout/late-writer errors
        # in per-property try blocks. Validation already happened for every key.
        for key,command in operations: lines += [command,f'set end of acceptedKeys to {_quote(key)}']
        if range_kind: lines += ['calculate obj']
        lines += ['return '+_object({'accepted':'acceptedKeys','rejected':'rejectedKeys'})]
        return self._run('\n'.join(lines))

    def apply_structural_op(self,op):
        self._assert_live(write=True)
        if not isinstance(op,dict): raise ValueError('Invalid Excel structural operation')
        verb=op.get('op');target=op.get('target');etype=op.get('type')
        if verb in ('move','clone'):
            _,kind,_=parse_target(target)
            if kind!='sheet': raise NotImplementedError('Excel range move/clone changes clipboard; no verified clipboard-free adapter')
            raise NotImplementedError('Excel worksheet move/clone session identity is not yet verified')
        if verb=='insert' and etype=='sheet':
            props=op.get('props') or {}
            if set(props)-{'name'}: raise ValueError('Unsupported worksheet insert fields')
            name=props.get('name','Sheet-'+uuid4().hex[:8]);quoted=_sheet_name(name)
            body=f'''if {quoted} is in (name of every worksheet of ownedBook) then error "Duplicate sheet name"
tell ownedBook
 make new worksheet at end with properties {{name:{quoted}}}
end tell
return {_object({'type':'"sheet"','path':'"sheet:" & (count of worksheets of ownedBook)'})}'''
            return self._run(body)
        if verb=='insert':
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
            if kind=='sheet': command='if (count of worksheets of ownedBook) <= 1 then error "Cannot remove last worksheet"\ndelete obj'
            elif kind in ('cell','range'):
                axis=op.get('axis','row')
                if axis not in ('row','column'): raise ValueError('Invalid Excel removal axis')
                command='delete range (entire '+axis+' of obj)'
            elif kind in ('shape','shape_name','chart'): command='delete obj'
            else: raise ValueError('Unsupported Excel removal target')
            return self._run(prefix+'\n'+command+'\nreturn '+_object({'removed':_quote(target)}))
        raise ValueError('Unsupported Excel structural operation')

    def _destination(self,path,suffix):
        destination=Path(path).expanduser().resolve()
        if destination==self._source: raise ValueError('File sessions never overwrite their source')
        if destination.suffix.lower()!=suffix: raise ValueError('Unexpected Excel output extension')
        if destination.exists(): raise FileExistsError(destination)
        return destination

    def save_current(self):
        self._assert_live(write=True)
        self._run('save ownedBook\nreturn "{}"')
        return str(self._native)

    def save(self,destination):
        self._assert_live()
        destination=self._destination(destination,'.xlsx')
        self._run('save ownedBook\nreturn "{}"')
        validate_before_deadline(ValidatorSpec.from_callable(validate_office_package,'xlsx'),self._native,self._deadline)
        destination.parent.mkdir(parents=True,exist_ok=True)
        copy_file_before_deadline(self._native,destination,deadline=self._deadline)
        return str(destination)

    def save_copy(self,destination): return self.save(destination)

    def export_pdf(self,destination):
        self._assert_live()
        destination=self._destination(destination,'.pdf')
        native_pdf=self._job/('export-'+uuid4().hex+'.pdf')
        self._run(f'save workbook as ownedBook filename {_quote(str(native_pdf))} file format PDF file format\nif full name of ownedBook is not {_quote(str(self._native))} then error "PDF export changed binding"\nreturn "{{}}"')
        validate_before_deadline(ValidatorSpec.from_callable(validate_pdf),native_pdf,self._deadline)
        destination.parent.mkdir(parents=True,exist_ok=True)
        copy_file_before_deadline(native_pdf,destination,deadline=self._deadline)
        return str(destination)

    def close(self,save_changes=False):
        if self._closed: return
        try:
            if not self._failed:
                if save_changes: self.save_current()
                self._run('close ownedBook saving no\nreturn "{}"')
        finally:
            self._closed=True
            if self._lock is not None: self._lock.close()
            # Keep scripts/evidence and the only edited recovery file; parent may
            # remove this private job explicitly after publication acceptance.
