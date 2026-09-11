"""Native, document-bound macOS PowerPoint inspection/edit sessions.

Files are staged in PowerPoint Data/Documents, never edited in place implicitly.
Unsupported dictionary primitives are explicit; failure retains recovery evidence.
"""
from __future__ import annotations

import json
import hashlib
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import time
from uuid import uuid4

from .macos_script import apple_string
from .macos_powerpoint_script import _rgb
from .macos_office_runtime import OfficeJobLock, _container_root, NativeOfficeError
from .input_validation import validate_native_input
from ..artifact_transport import snapshot_artifact_state, publish_artifact, validate_office_package, validate_pdf, copy_file_before_deadline, validate_before_deadline, ValidatorSpec


class PowerPointSessionCapabilityError(ValueError):
    pass


_JSON = '''use framework "Foundation"
use scripting additions
on normalizedJSON(v)
 if v is missing value then return "__NATIVE_UNAVAILABLE__"
 if class of v is list then
  set resultArray to current application's NSMutableArray's array()
  repeat with itemRef in v
   set normalizedItem to my normalizedJSON(contents of itemRef)
   resultArray's addObject:normalizedItem
  end repeat
  return resultArray
 end if
 if class of v is reference then return my normalizedJSON(contents of v)
 if class of v is integer or class of v is real or class of v is boolean or class of v is text then return v
 return v as text
end normalizedJSON
on encodeJSON(itemsList)
 set normalizedItems to my normalizedJSON(itemsList)
 set jsonData to current application's NSJSONSerialization's dataWithJSONObject:normalizedItems options:0 |error|:(missing value)
 if jsonData is missing value then error "Native snapshot JSON encoding failed"
 return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
end encodeJSON
'''


def _number(value, *, positive=False):
    if isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value) or (positive and value <= 0):
        raise ValueError('Expected a finite native measurement')
    return str(value)


def _boolean(value):
    if value not in (True, False, -1, 0, 1):
        raise ValueError('Expected a boolean')
    return 'true' if bool(value) else 'false'


def _color(value):
    if not (isinstance(value,str) and re.fullmatch(r'#[0-9a-fA-F]{6}',value)) and not (type(value)is int and 0<=value<=0xffffff):
        raise ValueError('Expected RGB hex or BGR integer')
    return _rgb(value)


def _enum(value, values):
    if isinstance(value,bool) or value not in values:
        raise ValueError('Unsupported native enum value: '+str(value))
    return values[value]


def _target(target):
    if target in {'presentation','selection'}:
        return '', target
    match=re.fullmatch(r'slide:([1-9]\d*)(?:/shape:([1-9]\d*|@id=\d+|@name=.+?))?(?:/paragraph:([1-9]\d*)(?:/run:([1-9]\d*))?|/table/cell:([1-9]\d*),([1-9]\d*))?',target)
    if not match:
        raise ValueError('Unsupported Slide target: '+str(target))
    si,sh,para,run,row,col=match.groups()
    lines=[f'set targetSlide to slide {si} of ownedDoc']
    if sh is None:
        return '\n'.join(lines),'slide'
    if sh.startswith('@id='):
        raise PowerPointSessionCapabilityError('Native PowerPoint dictionary does not expose a verified stable shape ID')
    if sh.startswith('@name='):
        lines += ['set matchingShapes to {}', 'repeat with shapeIndex from 1 to (count of shapes of targetSlide)',
                  f' if name of shape shapeIndex of targetSlide is {apple_string(sh[6:])} then set end of matchingShapes to shapeIndex',
                  'end repeat','if (count of matchingShapes) is not 1 then error "Stale or ambiguous shape name"',
                  'set targetShape to shape (item 1 of matchingShapes) of targetSlide']
    else:
        lines += [f'set targetShape to shape {sh} of targetSlide']
    if row:
        lines += [f'set targetCell to get cell from table object of targetShape row {row} column {col}', 'set targetShape to shape of targetCell']
    lines += ['set targetText to text range of text frame of targetShape']
    if para:
        lines += [f'set targetText to paragraph {para} of targetText']
    if run:
        lines += [f'set targetText to text flow {run} of targetText']
    return '\n'.join(lines), 'paragraph' if para else ('cell' if row else 'shape')


ALIGN={1:'paragraph align left',2:'paragraph align center',3:'paragraph align right',4:'paragraph align justify',5:'paragraph align distribute',6:'paragraph align Thai',7:'paragraph align justify low'}
ANCHOR={1:'anchor top',2:'anchor top baseline',3:'anchor middle',4:'anchor bottom',5:'anchor bottom baseline'}
AUTO={0:'auto size none',1:'shape to fit text',2:'text to fit shape'}
DASH={i:'line dash style '+v for i,v in enumerate(('solid','square dot','round dot','dash','dash dot','dash dot dot','long dash','long dash dot','long dash dot dot','system dash','system dot','system dash dot'),1)}
LAYOUT={1:'slide layout title slide',2:'slide layout text slide',3:'slide layout two column text',4:'slide layout table',5:'slide layout text and chart',6:'slide layout chart',7:'slide layout orgchart',8:'slide layout chart and text',9:'slide layout text and clipart',10:'slide layout clipart and text',11:'slide layout title only',12:'slide layout blank'}


def compile_patch(target, *, text=None, font=None, paragraph=None, geometry=None, fill=None, line=None, text_frame=None, name=None, shape_type=None, background=None, follow_master_background=None, page_setup=None, vertical_alignment=None):
    """Validate entire patch before emitting any document mutation."""
    binding,kind=_target(target)
    assignments=[];accepted=[];rejected=[]
    def add(label,prop,value,encoder=apple_string):
        encoded=encoder(value)
        assignments.append(f'set {prop} to {encoded}')
        accepted.append(label)
    if kind=='selection':
        binding='set boundSelection to selection of document window 1 of ownedDoc\nif selection type of boundSelection is not selection type text then error "Native selected-shape patch is not verified"\nset targetText to text range of boundSelection'
        kind='paragraph'
    if text is not None:
        if kind in {'shape','cell','paragraph'}:add('text','content of targetText',text)
        else:rejected.append('text')
    for key,value in (font or {}).items():
        prop={'name':'font name','size':'font size','bold':'bold','italic':'italic','underline':'underline','strikethrough':'strike type','color':'font color'}.get(key)
        if not prop or kind not in {'shape','cell','paragraph'}:rejected.append('font.'+key);continue
        encoder=_color if key=='color' else (lambda x:_number(x,positive=True)) if key=='size' else _boolean if key in {'bold','italic','underline'} else apple_string
        if key=='strikethrough':encoder=lambda x:'single strike' if _boolean(x)=='true' else 'no strike'
        add('font.'+key,prop+' of font of targetText',value,encoder)
        if key=='name':assignments.append('set east asian name of font of targetText to '+apple_string(value))
    for key,value in (paragraph or {}).items():
        prop={'alignment':'alignment','space_before':'space before','space_after':'space after','line_spacing':'space within'}.get(key)
        if not prop or kind not in {'shape','cell','paragraph'}:rejected.append('paragraph.'+key);continue
        add('paragraph.'+key,prop+' of paragraph format of targetText',value,(lambda x:_enum(x,ALIGN)) if key=='alignment' else _number)
    for key,value in (geometry or {}).items():
        prop={'left':'left position','top':'top','width':'width','height':'height','rotation':'rotation'}.get(key)
        if prop and kind=='shape':add('geometry.'+key,prop+' of targetShape',value,lambda x:_number(x,positive=key in {'width','height'}))
        else:rejected.append('geometry.'+key)
    for dimension,patch,obj in [('fill',fill,'fill format of targetShape'),('line',line,'line format of targetShape'),('background',background,'fill format of background of targetSlide')]:
        for key,value in (patch or {}).items():
            valid=kind in {'shape','cell'} if dimension!='background' else kind=='slide'
            prop={'color':'fore color','back_color':'back color','visible':'visible','transparency':'transparency','weight':'line weight','dash_style':'dash style'}.get(key)
            if not valid or not prop or (dimension!='line' and key in {'weight','dash_style'}) or (dimension=='line' and key in {'back_color','visible'}):
                rejected.append(dimension+'.'+key);continue
            encoder=_color if key in {'color','back_color'} else _boolean if key=='visible' else (lambda x:_enum(x,DASH)) if key=='dash_style' else _number
            if key=='transparency' and (isinstance(value,bool) or not isinstance(value,(int,float)) or not 0<=value<=1):raise ValueError('Transparency must be between 0 and 1')
            add(dimension+'.'+key,prop+' of '+obj,value,encoder)
    tf=dict(text_frame or {})
    if vertical_alignment is not None:tf['vertical_anchor']=vertical_alignment
    for key,value in tf.items():
        prop={'margin_left':'margin left','margin_right':'margin right','margin_top':'margin top','margin_bottom':'margin bottom','word_wrap':'word wrap','auto_size':'auto size','vertical_anchor':'vertical anchor'}.get(key)
        if not prop or kind not in {'shape','cell'}:rejected.append('text_frame.'+key);continue
        encoder=_boolean if key=='word_wrap' else (lambda x:_enum(x,AUTO)) if key=='auto_size' else (lambda x:_enum(x,ANCHOR)) if key=='vertical_anchor' else _number
        add('text_frame.'+key,prop+' of text frame of targetShape',value,encoder)
    if name is not None:
        if kind=='shape':add('name','name of targetShape',name)
        else:rejected.append('name')
    if shape_type is not None:rejected.append('shape_type')
    if follow_master_background is not None:
        if kind=='slide':add('follow_master_background','follow master background of targetSlide',follow_master_background,_boolean)
        else:rejected.append('follow_master_background')
    for key,value in (page_setup or {}).items():
        if kind=='presentation' and key=='slide_width':add('page_setup.'+key,'slide width of page setup of ownedDoc',value,lambda x:_number(x,positive=True))
        else:rejected.append('page_setup.'+key)
    if rejected:return '',{'accepted':[],'rejected':rejected}
    # Native toggles can reset dependent attributes. Final colors and bounds
    # must be assigned after those toggles, independent of caller dict order.
    def emission_order(statement):
        if statement.startswith('set follow master background'):return 0
        if any(statement.startswith('set '+prop+' of targetShape') for prop in ('left position','top','width','height','rotation')):
            return 6 if statement.startswith('set height') else 5
        if statement.startswith(('set fore color','set back color')):return 4
        if statement.startswith(('set line weight','set visible')):return 3
        return 2
    assignments.sort(key=emission_order)
    if kind in {'shape','cell'} and not any('targetText' in item for item in assignments):
        binding=binding.replace('\nset targetText to text range of text frame of targetShape','')
    return ('\n'.join([binding]+assignments) if assignments else ''), {'accepted':accepted,'rejected':rejected}


AUTOSHAPES={1: 'autoshape rectangle', 2: 'autoshape parallelogram', 3: 'autoshape trapezoid', 4: 'autoshape diamond', 5: 'autoshape rounded rectangle', 6: 'autoshape octagon', 7: 'autoshape isosceles triangle', 8: 'autoshape right triangle', 9: 'autoshape oval', 10: 'autoshape hexagon', 11: 'autoshape cross', 12: 'autoshape regular pentagon', 13: 'autoshape can', 14: 'autoshape cube', 15: 'autoshape bevel', 16: 'autoshape folded corner', 17: 'autoshape smiley face', 18: 'autoshape donut', 19: 'autoshape no symbol', 20: 'autoshape block arc', 21: 'autoshape heart', 22: 'autoshape lightning bolt', 23: 'autoshape sun', 24: 'autoshape moon', 25: 'autoshape arc', 26: 'autoshape double bracket', 27: 'autoshape double brace', 28: 'autoshape plaque', 29: 'autoshape left bracket', 30: 'autoshape right bracket', 31: 'autoshape left brace', 32: 'autoshape right brace', 33: 'autoshape right arrow', 34: 'autoshape left arrow', 35: 'autoshape up arrow', 36: 'autoshape down arrow', 37: 'autoshape left right arrow', 38: 'autoshape up down arrow', 39: 'autoshape quad arrow', 40: 'autoshape left right up arrow', 41: 'autoshape bent arrow', 42: 'autoshape U turn arrow', 43: 'autoshape left up arrow', 44: 'autoshape bent up arrow', 45: 'autoshape curved right arrow', 46: 'autoshape curved left arrow', 47: 'autoshape curved up arrow', 48: 'autoshape curved down arrow', 49: 'autoshape striped right arrow', 50: 'autoshape notched right arrow', 51: 'autoshape pentagon', 52: 'autoshape chevron', 53: 'autoshape right arrow callout', 54: 'autoshape left arrow callout', 55: 'autoshape up arrow callout', 56: 'autoshape down arrow callout', 57: 'autoshape left right arrow callout', 58: 'autoshape up down arrow callout', 59: 'autoshape quad arrow callout', 60: 'autoshape circular arrow', 61: 'autoshape flowchart process', 62: 'autoshape flowchart alternate process', 63: 'autoshape flowchart decision', 64: 'autoshape flowchart data', 65: 'autoshape flowchart predefined process', 66: 'autoshape flowchart internal storage', 67: 'autoshape flowchart document', 68: 'autoshape flowchart multi document', 69: 'autoshape flowchart terminator', 70: 'autoshape flowchart preparation', 71: 'autoshape flowchart manual input', 72: 'autoshape flowchart manual operation', 73: 'autoshape flowchart connector', 74: 'autoshape flowchart offpage connector', 75: 'autoshape flowchart card', 76: 'autoshape flowchart punched tape', 77: 'autoshape flowchart summing junction', 78: 'autoshape flowchart or', 79: 'autoshape flowchart collate', 80: 'autoshape flowchart sort', 81: 'autoshape flowchart extract', 82: 'autoshape flowchart merge', 83: 'autoshape flowchart stored data', 84: 'autoshape flowchart delay', 85: 'autoshape flowchart sequential access storage', 86: 'autoshape flowchart magnetic disk', 87: 'autoshape flowchart direct access storage', 88: 'autoshape flowchart display', 89: 'autoshape explosion one', 90: 'autoshape explosion two', 91: 'autoshape four point star', 92: 'autoshape five point star', 93: 'autoshape eight point star', 94: 'autoshape sixteen point star', 95: 'autoshape twenty four point star', 96: 'autoshape thirty two point star', 97: 'autoshape up ribbon', 98: 'autoshape down ribbon', 99: 'autoshape curved up ribbon', 100: 'autoshape curved down ribbon', 101: 'autoshape vertical scroll', 102: 'autoshape horizontal scroll', 103: 'autoshape wave', 104: 'autoshape double wave', 105: 'autoshape rectangular callout', 106: 'autoshape rounded rectangular callout', 107: 'autoshape oval callout', 108: 'autoshape cloud callout', 109: 'autoshape line callout one', 110: 'autoshape line callout two', 111: 'autoshape line callout three', 112: 'autoshape line callout four', 113: 'autoshape line callout one accent bar', 114: 'autoshape line callout two accent bar', 115: 'autoshape line callout three accent bar', 116: 'autoshape line callout four accent bar', 117: 'autoshape line callout one no border', 118: 'autoshape line callout two no border', 119: 'autoshape line callout three no border', 120: 'autoshape line callout four no border', 121: 'autoshape callout one border and accent bar', 122: 'autoshape callout two border and accent bar', 123: 'autoshape callout three border and accent bar', 124: 'autoshape callout four border and accent bar', 125: 'autoshape action button custom', 126: 'autoshape action button home', 127: 'autoshape action button help', 128: 'autoshape action button information', 129: 'autoshape action button back or previous', 130: 'autoshape action button forward or next', 131: 'autoshape action button beginning', 132: 'autoshape action button end', 133: 'autoshape action button return', 134: 'autoshape action button document', 135: 'autoshape action button sound', 136: 'autoshape action button movie', 137: 'autoshape balloon', 138: 'autoshape not primitive', 139: 'autoshape flowchart offline storage', 140: 'autoshape left right ribbon', 141: 'autoshape diagonal stripe', 142: 'autoshape pie', 143: 'autoshape non isosceles trapezoid', 144: 'autoshape Decagon', 145: 'autoshape Heptagon', 146: 'autoshape Dodecagon', 147: 'autoshape six points star', 148: 'autoshape seven points star', 149: 'autoshape ten points star', 150: 'autoshape twelve points star', 151: 'autoshape round one rectangle', 152: 'autoshape round two same rectangle', 153: 'autoshape round two diagonal rectangle', 154: 'autoshape snip round rectangle', 155: 'autoshape snip one rectangle', 156: 'autoshape snip two same rectangle', 157: 'autoshape snip two diagonal rectangle', 158: 'autoshape frame', 159: 'autoshape half frame', 160: 'autoshape tear', 161: 'autoshape chord', 162: 'autoshape corner', 163: 'autoshape math plus', 164: 'autoshape math minus', 165: 'autoshape math multiply', 166: 'autoshape math divide', 167: 'autoshape math equal', 168: 'autoshape math not equal', 169: 'autoshape corner tabs', 170: 'autoshape square tabs', 171: 'autoshape plaque tabs', 172: 'autoshape gear six', 173: 'autoshape gear nine', 174: 'autoshape funnel', 175: 'autoshape pie wedge', 176: 'autoshape left circular arrow', 177: 'autoshape left right circular arrow', 178: 'autoshape swoosh arrow', 179: 'autoshape cloud', 180: 'autoshape chart x', 181: 'autoshape chart star', 182: 'autoshape chart plus', 183: 'autoshape line inverse'}

def _index(value):
    if type(value)is not int or value<1:raise ValueError("Invalid native slide index")
    return value


SHAPE_TYPES={'shape type unset': 65534, 'shape type auto': 1, 'shape type callout': 2, 'shape type chart': 3, 'shape type comment': 4, 'shape type free form': 5, 'shape type group': 6, 'shape type embedded OLE control': 7, 'shape type form control': 8, 'shape type line': 9, 'shape type linked OLE object': 10, 'shape type linked picture': 11, 'shape type OLE control': 12, 'shape type picture': 13, 'shape type place holder': 14, 'shape type word art': 15, 'shape type media': 16, 'shape type text box': 17, 'shape type script anchor': 18, 'shape type table': 19, 'shape type canvas': 20, 'shape type diagram': 21, 'shape type ink': 22, 'shape type ink comment': 23, 'shape type smartart graphic': 24, 'shape type slicer': 25, 'shape type web video': 26, 'shape type content application': 27, 'shape type graphic': 28, 'shape type linked graphic': 29, 'shape type 3d model': 30, 'shape type linked 3d model': 31}

def _native_enum(value,mapping):
    return {name:number for number,name in mapping.items()}.get(value)


def _hex(value):
    return '#'+''.join(f'{int(v):02X}' for v in value) if isinstance(value,list) and len(value)==3 else None


def parse_snapshot(rows, *, include_text=True, max_shapes=None):
    def clean(value):
        if isinstance(value,list):return [clean(v) for v in value]
        return None if value == '__NATIVE_UNAVAILABLE__' or value == -2147483648 else value
    rows=clean(rows)
    doc=next(r for r in rows if r[0]=='doc')
    result={'kind':'slide','engine':'msoffice','name':doc[1],'path':doc[2],'saved':doc[3],'slide_count':doc[4],'page_setup':{'slide_width':doc[5]},'shapes_truncated':False,'slides':[],'limitations':['Native numeric Shape.Id is not exposed by this adapter','Independent page height unavailable in native dictionary']}
    slides={};shapes={};name_counts={};remaining=max_shapes if max_shapes is not None else math.inf
    for row in rows:
        if row[0]=='shape':name_counts[(row[1],row[3])]=name_counts.get((row[1],row[3]),0)+1
    for r in rows:
        if r[0]=='slide':
            slide={'id':f'slide:{r[1]}','index':r[1],'native_slide_id':r[2],'name':None,'layout':_native_enum(r[3],LAYOUT),'native_layout':r[3],'notes':r[4] if include_text else None,'follow_master_background':r[5],'background':{'color':_hex(r[6])},'shape_count':0,'shapes':[]};slides[r[1]]=slide;result['slides'].append(slide)
        elif r[0]=='shape':
            slide=slides[r[1]];slide['shape_count']+=1
            if remaining<=0:result['shapes_truncated']=True;continue
            remaining-=1
            shape={'id':f'slide:{r[1]}/shape:@name={r[3]}','index':r[2],'shape_id':None,'name':r[3],'type':SHAPE_TYPES.get(r[4]),'native_type':r[4],'geometry':dict(zip(('left','top','width','height','rotation','z_order'),r[5:11])),'fill':{'color':_hex(r[11]),'visible':r[12],'transparency':r[13]},'font':dict(zip(('name','size','bold','italic','underline'),r[15:20])),'paragraphs':[]}
            shape['font']['color']=_hex(r[20])
            if include_text:shape['text']=r[14]
            shapes[(r[1],r[2])]=shape;slide['shapes'].append(shape)
        elif r[0]=='paragraph' and (r[1],r[2]) in shapes:
            shape=shapes[(r[1],r[2])];entry={'id':shape['id']+f'/paragraph:{r[3]}','index':r[3],'runs':[]}
            if include_text:entry['text']=r[4]
            shape['paragraphs'].append(entry)
        elif r[0]=='run' and (r[1],r[2]) in shapes:
            shape=shapes[(r[1],r[2])]
            para=next(p for p in shape['paragraphs'] if p['index']==r[3])
            entry={'id':para['id']+f'/run:{r[4]}','index':r[4],'font':dict(zip(('name','size','bold','italic','underline'),r[6:11]))}
            entry['font']['color']=_hex(r[11]);entry['font']['strikethrough']={'no strike':False,'single strike':True,'double strike':True}.get(r[12])
            if include_text:entry['text']=r[5]
            para['runs'].append(entry)
        elif r[0]=='cell' and (r[1],r[2]) in shapes:
            shape=shapes[(r[1],r[2])];table=shape.setdefault('table',{'rows':r[3],'columns':r[4],'cells':[]});cell={'id':shape['id']+f'/table/cell:{r[5]},{r[6]}','row':r[5],'column':r[6]}
            if include_text:cell['text']=r[7]
            if len(r)>8:
                cell['font']=dict(zip(('name','size','bold','italic','underline'),r[8:13]));cell['font']['color']=_hex(r[13])
                cell['paragraph']={'alignment':_native_enum(r[14],ALIGN),'space_before':r[15],'space_after':r[16],'line_spacing':r[17]}
                cell['fill']={'color':_hex(r[18])}
            table['cells'].append(cell)
        elif r[0]=='detail' and (r[1],r[2]) in shapes:
            shapes[(r[1],r[2])]['line']={'color':_hex(r[3]),'weight':r[4],'visible':r[5],'transparency':r[6]}
            shapes[(r[1],r[2])]['text_frame']=dict(zip(('margin_left','margin_right','margin_top','margin_bottom','word_wrap'),r[7:12]))
            shapes[(r[1],r[2])]['paragraph']={'alignment':_native_enum(r[12],ALIGN),'space_before':r[13],'space_after':r[14],'line_spacing':r[15]}
    # Name-based addressing is stable only when unique. Never silently pick one.
    for slide in result['slides']:
        for shape in slide['shapes']:
            if name_counts[(slide['index'],shape['name'])]>1:
                previous=shape['id']
                shape['id']=f"slide:{slide['index']}/shape:{shape['index']}"
                descendants=shape['paragraphs']+shape.get('table',{}).get('cells',[])
                for entry in descendants:
                    entry['id']=shape['id']+entry['id'][len(previous):]
                    for run in entry.get('runs',[]):run['id']=shape['id']+run['id'][len(previous):]
                shape['identity_warning']='Duplicate shape name; positional target only'
    return result


_SNAPSHOT='''set snapshotRows to {{"doc", name of ownedDoc, full name of ownedDoc, saved of ownedDoc, count of slides of ownedDoc, slide width of page setup of ownedDoc}}
repeat with si from 1 to (count of slides of ownedDoc)
 set sl to slide si of ownedDoc
 set noteText to ""
 try
  set noteText to content of text range of text frame of shape 2 of notes page of sl
 end try
 set end of snapshotRows to {"slide", si as integer, slide ID of sl, layout of sl as text, noteText, follow master background of sl, fore color of fill format of background of sl}
 repeat with sh from 1 to (count of shapes of sl)
  set shp to shape sh of sl
  set txt to ""
  set fn to ""
  set fs to 0
  set fb to false
  set fi to false
  set fu to false
  set fc to {0,0,0}
  if has text frame of shp then
   set tx to text range of text frame of shp
   set txt to content of tx
   set fn to font name of font of tx
   set fs to font size of font of tx
   set fb to bold of font of tx
   set fi to italic of font of tx
   set fu to underline of font of tx
   set fc to font color of font of tx
  end if
  set end of snapshotRows to {"shape",si as integer,sh as integer,name of shp,shape type of shp as text,left position of shp,top of shp,width of shp,height of shp,rotation of shp,z order position of shp,fore color of fill format of shp,visible of fill format of shp,transparency of fill format of shp,txt,fn,fs,fb,fi,fu,fc}
  if has text frame of shp then
   set tx to text range of text frame of shp
   repeat with pi from 1 to (count of paragraphs of tx)
    set paraText to paragraph pi of tx
    set end of snapshotRows to {"paragraph",si as integer,sh as integer,pi as integer,content of paraText}
    repeat with runIndex from 1 to (count of text flows of paraText)
     set runText to text flow runIndex of paraText
     set runFont to font of runText
     set end of snapshotRows to {"run",si as integer,sh as integer,pi as integer,runIndex as integer,content of runText,font name of runFont,font size of runFont,bold of runFont,italic of runFont,underline of runFont,font color of runFont,strike type of runFont as text}
    end repeat
   end repeat
   set pf to paragraph format of tx
   set tf to text frame of shp
   set ln to line format of shp
   set end of snapshotRows to {"detail",si as integer,sh as integer,fore color of ln,line weight of ln,"__NATIVE_UNAVAILABLE__",transparency of ln,margin left of tf,margin right of tf,margin top of tf,margin bottom of tf,word wrap of tf,alignment of pf as text,space before of pf,space after of pf,space within of pf}
  end if
  if has table of shp then
   set nr to number of rows of shp
   set nc to number of columns of shp
   repeat with ri from 1 to nr
    repeat with ci from 1 to nc
     set cc to get cell from table object of shp row ri column ci
     set cellShape to shape of cc
     set cellText to text range of text frame of cellShape
     set cellFont to font of cellText
     set cellParagraph to paragraph format of cellText
     set end of snapshotRows to {"cell",si as integer,sh as integer,nr,nc,ri as integer,ci as integer,content of cellText,font name of cellFont,font size of cellFont,bold of cellFont,italic of cellFont,underline of cellFont,font color of cellFont,alignment of cellParagraph as text,space before of cellParagraph,space after of cellParagraph,space within of cellParagraph,fore color of fill format of cellShape}
    end repeat
   end repeat
  end if
 end repeat
end repeat
return my encodeJSON(snapshotRows)'''


class MacPowerPointSession:
    kind='slide'
    engine='msoffice'
    supports_attached_save_copy=False

    def __init__(self,path=None,read_only=False,visible=False,*,timeout=45):
        self._source=Path(path).expanduser().resolve() if path else None
        self._read_only=read_only;self._visible=visible;self.timeout=float(timeout)
        if not math.isfinite(self.timeout) or self.timeout<=0:raise ValueError('Invalid timeout')
        self._native_read_only=False;self._deadline=None;self._entered=False;self._attached=False;self._lock=None;self._job=None;self._path=None;self._name=None;self._uncertain=False;self._sequence=0
        self._source_digest=None;self._bound_digest=None
        self._logical=self._source;self._logical_state=None

    @classmethod
    def new_document(cls,visible=False):
        return cls(visible=visible)

    @classmethod
    def open_document(cls,path,*,read_only=False,visible=False):
        return cls(path,read_only,visible)

    @classmethod
    def attach_active(cls):
        session=cls();session._attached=True
        session.__enter__()
        return session

    def _remaining(self):
        remaining=self._deadline-time.monotonic()
        if remaining<=0:
            if self._entered:self._quarantine('Session deadline expired before confirmed native close')
            raise NativeOfficeError('NATIVE_OFFICE_TIMEOUT',staging_path=self._job)
        return remaining

    def _prepare(self):
        if sys.platform!='darwin':raise RuntimeError('PowerPoint session requires macOS')
        root=_container_root('presentation');root.mkdir(parents=True,exist_ok=True)
        self._lock=OfficeJobLock(root);self._lock.acquire(self._deadline)
        self._job=Path(tempfile.mkdtemp(prefix='session-',dir=root))

    def _digest(self,path):
        digest=hashlib.sha256()
        with Path(path).open('rb') as stream:
            while True:
                self._remaining()
                chunk=stream.read(1024*1024)
                if not chunk:break
                digest.update(chunk)
        return digest.hexdigest()

    def _verify_source_digest(self):
        path=self._path if self._attached else self._source
        expected=self._bound_digest if self._attached else self._source_digest
        if path is None or expected is None:return
        try:actual=self._digest(path)
        except OSError:raise ValueError('Source file changed during PowerPoint session') from None
        if actual!=expected:raise ValueError('Source file changed during PowerPoint session')

    def __enter__(self):
        if self._entered:return self
        self._deadline=time.monotonic()+self.timeout
        if self._source is not None:validate_native_input(self._source,'presentation',deadline=self._deadline)
        try:
            self._prepare()
            if self._attached:
                body='''set activeDoc to active presentation
set activeName to name of activeDoc
set activePath to full name of activeDoc
set matched to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is activeName then set matched to matched + 1
end repeat
if matched is not 1 then error "Ambiguous active presentation"
return my encodeJSON({activeName,activePath,saved of activeDoc,read only of activeDoc})'''
                identity=json.loads(self._run(body,bound=False,prefix='if application "Microsoft PowerPoint" is not running then error "PowerPoint is not running"\n'));self._name=identity[0];self._path=identity[1];self._native_read_only=identity[3]
                if Path(str(self._path)).is_absolute() and Path(self._path).is_file():self._bound_digest=self._digest(self._path)
            else:
                self._path=str(self._job/('bound-'+uuid4().hex+'.pptx'));self._name=Path(self._path).name
                if self._source is not None:
                    self._source_digest=self._digest(self._source)
                    self._logical_state=snapshot_artifact_state(self._source,deadline=self._deadline)
                    copy_file_before_deadline(self._source,Path(self._path),deadline=self._deadline)
                    validate_native_input(Path(self._path),'presentation',deadline=self._deadline)
                    if self._digest(self._path)!=self._source_digest:raise ValueError('Source file changed during PowerPoint staging')
                    if snapshot_artifact_state(self._source,deadline=self._deadline)!=self._logical_state:raise ValueError('Source file changed during PowerPoint staging')
                if self._source:
                    prefix=f'set inputFile to (POSIX file {apple_string(self._path)}) as alias\n'
                    body=f'if exists presentation {apple_string(self._name)} then error "Presentation collision"\nopen inputFile\nset ownedDoc to presentation {apple_string(self._name)}\nif full name of ownedDoc is not {apple_string(self._path)} then error "Owned path mismatch"\nreturn "OPEN"'
                    self._run(body,bound=False,prefix=prefix,mutation=True)
                else:
                    prefix=f'set nativeHFS to (POSIX file {apple_string(self._path)}) as text\n'
                    body=f'''if exists presentation {apple_string(self._name)} then error "Presentation collision"
set ownedDoc to make new presentation
log "OWNED_UNSAVED_NAME|" & name of ownedDoc
save ownedDoc in nativeHFS as save as Open XML presentation
set ownedDoc to presentation {apple_string(self._name)}
if full name of ownedDoc is not {apple_string(self._path)} then error "Owned path mismatch"
return "CREATED"'''
                    self._run(body,bound=False,prefix=prefix,mutation=True)
            self._entered=True
            return self
        except BaseException:
            self._release()
            raise

    def _persist_diagnostic(self,label,write):
        try:
            write()
        except BaseException as exc:
            if not hasattr(self,'_diagnostic_io_failures'):self._diagnostic_io_failures=[]
            self._diagnostic_io_failures.append((label,type(exc).__name__))
            return exc
        return None

    def _quarantine(self,reason):
        self._uncertain=True
        detail={'component':'presentation','staging_path':str(self._job),'bound_path':str(self._path),'reason':reason}
        if self._job is not None:
            self._persist_diagnostic('recovery',lambda:(self._job/'recovery.json').write_text(json.dumps(detail,ensure_ascii=False)))
        if self._lock:
            def persist_lock():
                if not self._lock.quarantine_path.exists():self._lock.quarantine(detail)
            self._persist_diagnostic('quarantine',persist_lock)

    def _write_logs(self,path,stdout,stderr):
        first_error=None;written={}
        for name,text in [('stdout',stdout),('stderr',stderr)]:
            error=self._persist_diagnostic(name,lambda name=name,text=text:
                path.with_suffix('.'+name).write_bytes(text if isinstance(text,bytes) else (text or '').encode()))
            written[name]=error is None
            if first_error is None:first_error=error
        return first_error,written

    def _run(self,body,*,bound=True,prefix='',mutation=False):
        if self._uncertain:raise RuntimeError('Native session is quarantined; inspect retained recovery evidence')
        if bound and not self._entered:raise RuntimeError('Session must be entered')
        guard=''
        if bound:
            guard=f'''set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is {apple_string(self._name)} then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation {apple_string(self._name)}
if full name of ownedDoc is not {apple_string(str(self._path))} then error "Bound presentation path changed"
'''
        script=_JSON+'on sessionOperation()\n'+prefix+f'with timeout of {max(1,int(self._remaining()))} seconds\n tell application "Microsoft PowerPoint"\n{guard}{body}\n end tell\nend timeout\nend sessionOperation\nset operationResult to my sessionOperation()\nreturn my encodeJSON({{"WPSCOMPOSER_PPT_SESSION_OK",operationResult}})\n'
        self._sequence+=1;path=self._job/f'{self._sequence:03}.applescript';path.write_text(script)
        budget=self._remaining()
        stdout=stderr=''
        try:
            result=subprocess.run(['/usr/bin/osascript',str(path)],capture_output=True,text=True,timeout=budget)
            stdout,stderr=result.stdout,result.stderr
            if result.returncode:raise NativeOfficeError('NATIVE_OFFICE_EXECUTION_FAILED',staging_path=self._job)
            try:envelope=json.loads(result.stdout)
            except (ValueError,TypeError):envelope=None
            if not isinstance(envelope,list) or len(envelope)!=2 or envelope[0]!='WPSCOMPOSER_PPT_SESSION_OK':
                raise NativeOfficeError('NATIVE_OFFICE_EXECUTION_FAILED',staging_path=self._job)
            output=str(envelope[1])
        except BaseException as exc:
            if isinstance(exc,subprocess.TimeoutExpired):stdout,stderr=exc.stdout or '',exc.stderr or ''
            # Classify native failure first; auxiliary persistence must never
            # replace the primary timeout, native error, or cancellation.
            if isinstance(exc,subprocess.TimeoutExpired) or mutation or not isinstance(exc,Exception):
                self._quarantine(type(exc).__name__)
            log_error,written=self._write_logs(path,stdout,stderr)
            if log_error is not None and not self._uncertain:
                self._quarantine('Native diagnostic persistence failed')
            if isinstance(exc,subprocess.TimeoutExpired):
                exc=NativeOfficeError('NATIVE_OFFICE_TIMEOUT',staging_path=self._job)
            if isinstance(exc,NativeOfficeError):
                diagnostic='stderr' if stderr else 'stdout'
                exc.diagnostic_path=str(path.with_suffix('.'+diagnostic)) if written[diagnostic] else None
                exc.diagnostic_io_failures=tuple(getattr(self,'_diagnostic_io_failures',()))
                if exc.diagnostic_io_failures:
                    exc.safe_message=('Native Office timed out; recovery required.' if exc.code=='NATIVE_OFFICE_TIMEOUT'
                                      else 'Native Office execution failed; recovery required.')
                    exc.args=(exc.safe_message,)
                raise exc from None
            raise
        log_error,_=self._write_logs(path,stdout,stderr)
        if log_error is not None:
            self._quarantine('Native diagnostic persistence failed')
            raise log_error
        return output

    def __exit__(self,*exc):self.close(save_changes=False)

    def _mutable(self):
        if self._read_only or self._native_read_only:raise PermissionError('Read-only PowerPoint session')

    def inspect_document(self,include_text=True,max_shapes=None):
        if max_shapes is not None and (type(max_shapes)is not int or max_shapes<0):raise ValueError('max_shapes must be nonnegative')
        return parse_snapshot(json.loads(self._run(_SNAPSHOT)),include_text=include_text,max_shapes=max_shapes)

    def inspect_selection(self):
        rows=json.loads(self._run('''if full name of active presentation is not full name of ownedDoc then error "Selection belongs to a different presentation"
set sel to selection of document window 1 of ownedDoc
if selection type of sel is selection type text then
 set tx to text range of sel
 set fn to font of tx
 set pf to paragraph format of tx
 return my encodeJSON({"text",content of tx,font name of fn,font size of fn,bold of fn,italic of fn,underline of fn,font color of fn,alignment of pf as text,space before of pf,space after of pf,space within of pf})
end if
if selection type of sel is selection type none then return my encodeJSON({"none"})
if selection type of sel is selection type shapes then
 if has child shape range of sel then error "Child group shape selection is unsupported"
 set selectedIndices to {}
 set selectedShapes to shape range of sel
 repeat with selectedIndex from 1 to count of shapes of selectedShapes
  set end of selectedIndices to z order position of shape selectedIndex of selectedShapes
 end repeat
 return my encodeJSON({"shapes",slide index of slide of view of document window 1 of ownedDoc,selectedIndices})
end if
return my encodeJSON({"unverified",selection type of sel as text})'''))
        if rows[0]=='text':
            font=dict(zip(('name','size','bold','italic','underline'),rows[2:7]));font['color']=_hex(rows[7])
            paragraph={'alignment':_native_enum(rows[8],ALIGN),'space_before':rows[9],'space_after':rows[10],'line_spacing':rows[11]}
            return {'kind':'slide_text_selection','text':rows[1],'font':font,'paragraph':paragraph}
        if rows[0]=='none':return {'kind':'slide_selection','count':0,'shapes':[]}
        if rows[0]=='shapes':
            slide=next(sl for sl in self.inspect_document()['slides'] if sl['index']==rows[1])
            shapes=[sh for sh in slide['shapes'] if sh['index'] in rows[2]]
            if len(shapes)!=len(rows[2]):raise PowerPointSessionCapabilityError('Selection changed during snapshot')
            return {'kind':'slide_selection','count':len(shapes),'shapes':shapes}
        raise PowerPointSessionCapabilityError('Native non-text selection inspection not yet verified: '+rows[1])

    def apply_format_patch(self,target,**patch):
        self._mutable()
        selection_guard=''
        if target=='selection':
            selection=self.inspect_selection()
            if selection['kind']=='slide_selection':
                if selection['count']!=1:raise PowerPointSessionCapabilityError('Shape-selection patch requires exactly one native shape; multi-object atomic mutation is unavailable')
                selected=selection['shapes'][0]
                target=selected['id']
                slide_index=int(re.match(r'slide:(\d+)',target).group(1))
                selection_guard=f'''if full name of active presentation is not full name of ownedDoc then error "Selection changed before mutation"
set boundSelection to selection of document window 1 of ownedDoc
if selection type of boundSelection is not selection type shapes then error "Selection changed before mutation"
if count of shapes of shape range of boundSelection is not 1 then error "Selection changed before mutation"
if slide index of slide of view of document window 1 of ownedDoc is not {slide_index} then error "Selection changed before mutation"
if z order position of shape 1 of shape range of boundSelection is not {selected['index']} then error "Selection changed before mutation"
'''
            else:
                selection_guard='if full name of active presentation is not full name of ownedDoc then error "Selection changed before mutation"\n'
        body,result=compile_patch(target,**patch)
        if body:body=selection_guard+body
        if body:self._run(body+'\nreturn "PATCHED"',mutation=True)
        return result

    def apply_structural_op(self,op):
        self._mutable()
        from ..document_api import validate_op
        validation=validate_op(op,'slide')
        if not validation.get('valid',False):raise ValueError(str(validation))
        from .edit_preflight import validate_powerpoint_structural
        validate_powerpoint_structural(op)
        verb=op.get('op');props=op.get('props') or {};target=op.get('target','');position=op.get('position','end')
        if verb=='insert':
            etype=op['type'];parent=op.get('parent')
            geo={k:props.get(k,d) for k,d in [('left',100),('top',100),('width',300),('height',100)]}
            for k,v in geo.items():_number(v,positive=k in {'width','height'})
            if etype=='slide':
                layout=_enum(props.get('layout',12),LAYOUT)
                if position not in (None,'start','end') and not isinstance(position,dict):raise ValueError('Invalid slide position')
                loc='end of ownedDoc' if position in (None,'end') else 'before slide 1 of ownedDoc' if position=='start' else self._position(position)
                command=f'set addedSlide to make new slide at {loc} with properties {{layout:{layout}}}'
                if isinstance(position,dict) and 'index' in position:
                    value=position['index']
                    command=f'if {value} is (count of slides of ownedDoc) + 1 then\nset addedSlide to make new slide at end of ownedDoc with properties {{layout:{layout}}}\nelse\n'+command+'\nend if'
                result=self._run(command+'\nreturn slide index of addedSlide',mutation=True)
                return {'type':'slide','path':'slide:'+str(int(result))}
            binding,kind=_target(parent or '')
            if kind!='slide':raise ValueError('Shape insert requires slide parent')
            if etype not in {'textbox','shape','image'}:raise PowerPointSessionCapabilityError('Unsupported insert type')
            extra=''
            if etype=='image':
                resource=Path(props.get('path','')).expanduser().resolve()
                if not resource.is_file() or resource.suffix.lower() not in {'.png','.jpg','.jpeg','.gif','.bmp','.tiff','.tif'}:raise ValueError('Invalid image resource')
                staged=self._job/('image-'+uuid4().hex+resource.suffix);copy_file_before_deadline(resource,staged,deadline=self._deadline)
                extra=', file name:'+apple_string(str(staged))+', link to file:false, save with document:true'
            native='picture' if etype=='image' else 'text box'
            body=binding+f'\nset addedShape to make new {native} at end of targetSlide with properties {{left position:{geo["left"]},top:{geo["top"]},width:{geo["width"]},height:{geo["height"]}{extra}}}'
            if 'text' in props:body+='\nset content of text range of text frame of addedShape to '+apple_string(props['text'])
            result=self._run(body+'\nreturn count of shapes of targetSlide',mutation=True)
            return {'type':etype,'path':parent+'/shape:'+str(int(result))}
        binding,kind=_target(target)
        binding=binding.replace('\nset targetText to text range of text frame of targetShape','')
        obj='targetSlide' if kind=='slide' else 'targetShape'
        if kind not in {'slide','shape'}:raise ValueError('Structural target must be slide or shape')
        if verb=='remove':self._run(binding+'\ndelete '+obj+'\nreturn "REMOVED"',mutation=True);return {'removed':target}
        if verb=='clone' and kind=='slide':
            movement=self._slide_movement(op.get('to','end'),obj='targetSlide')
            body=binding+"""
set originalSlideIds to {}
repeat with nativeIndex from 1 to count of slides of ownedDoc
 set end of originalSlideIds to slide ID of slide nativeIndex of ownedDoc
end repeat
copy object targetSlide
paste object ownedDoc
set copiedIndices to {}
repeat with nativeIndex from 1 to count of slides of ownedDoc
 if slide ID of slide nativeIndex of ownedDoc is not in originalSlideIds then set end of copiedIndices to nativeIndex as integer
end repeat
if count of copiedIndices is not 1 then error "Native slide paste count mismatch"
set targetSlide to slide (item 1 of copiedIndices) of ownedDoc
"""+movement+'\nreturn slide index of targetSlide'
            index=self._clipboard_run(body)
            return {'type':'slide','cloned':True,'from':target,'path':'slide:'+str(int(index)),'clipboard_changed':True}
        if kind=='shape':
            si=int(re.match(r'slide:(\d+)',target).group(1));dest=op.get('to')
            destination=dest.get('slide',si) if isinstance(dest,dict) else si
            if type(destination)is not int or destination<1:raise ValueError('Invalid destination slide')
            if verb=='clone' and destination==si:
                self._run(binding+'\nset copiedShape to duplicate targetShape\nreturn count of shapes of targetSlide',mutation=True)
                return {'type':'shape','cloned':True,'from':target,'to_slide':si,'clipboard_changed':False}
            if verb in {'move','clone'}:
                command='copy shape'
                body=binding+f"""
set destinationSlide to slide {destination} of ownedDoc
set destinationCount to count of shapes of destinationSlide
set sourceGeometry to {{left position of targetShape,top of targetShape,width of targetShape,height of targetShape,rotation of targetShape}}
set sourceName to name of targetShape
set boundView to view of document window 1 of ownedDoc
set slide of boundView to destinationSlide
{command} targetShape
paste object boundView
if (count of shapes of destinationSlide) is not (destinationCount + 1) then error "Native shape paste count mismatch"
"""
                if verb=='move':
                    body+='''
set pastedShape to last shape of destinationSlide
set left position of pastedShape to item 1 of sourceGeometry
set top of pastedShape to item 2 of sourceGeometry
set width of pastedShape to item 3 of sourceGeometry
set height of pastedShape to item 4 of sourceGeometry
set rotation of pastedShape to item 5 of sourceGeometry
delete targetShape
set name of last shape of destinationSlide to sourceName
'''
                body+='\nreturn count of shapes of destinationSlide'
                index=self._clipboard_run(body)
                return {'type':'shape','moved' if verb=='move' else 'cloned':True,'from':target,'to_slide':destination,'path':f'slide:{destination}/shape:{int(index)}','clipboard_changed':True}
        if verb=='move':
            body=binding+'\n'+self._slide_movement(op.get('to','end'))+'\nreturn "MOVED"'
            self._run(body,mutation=True)
            return {'type':'slide','moved':True,'from':target,'clipboard_changed':False}
        raise ValueError('Unsupported structural verb')

    def _clipboard_run(self,body):
        try:return self._run(body,mutation=True)
        except BaseException as exc:
            # The native clipboard command may have completed before failure.
            exc.clipboard_changed=True
            raise

    @staticmethod
    def _slide_movement(to,obj='targetSlide'):
        if to in (None,'end'):expression='count of slides of ownedDoc'
        elif to=='start':expression='1'
        elif isinstance(to,dict) and 'index' in to:
            value=to['index']
            if type(value)is not int or value<1:raise ValueError('Invalid slide index')
            expression=str(value)
        elif isinstance(to,dict) and ('before' in to or 'after' in to):
            key='before' if 'before' in to else 'after'
            m=re.fullmatch(r'slide:([1-9]\d*)',str(to[key]))
            if not m:raise ValueError('Invalid slide anchor')
            return f'if slide index of {obj} is not {m[1]} then move {obj} to {key} slide {m[1]} of ownedDoc'
        else:raise ValueError('Invalid slide position')
        return f"""set destinationIndex to {expression}
if destinationIndex > count of slides of ownedDoc then set destinationIndex to count of slides of ownedDoc
set sourceIndex to slide index of {obj}
if sourceIndex < destinationIndex then
 move {obj} to after slide destinationIndex of ownedDoc
else if sourceIndex > destinationIndex then
 move {obj} to before slide destinationIndex of ownedDoc
end if"""

    @staticmethod
    def _position(position):
        if not isinstance(position,dict):raise ValueError('Invalid slide position')
        for key in ('before','after'):
            if key in position:
                m=re.fullmatch(r'slide:([1-9]\d*)',str(position[key]))
                if not m:raise ValueError('Invalid slide anchor')
                return f'{key} slide {m.group(1)} of ownedDoc'
        value=position.get('index')
        if type(value)is not int or value<1:raise ValueError('Invalid slide index')
        return f'before slide {value} of ownedDoc'

    LAYOUT_TITLE=1
    LAYOUT_TITLE_CONTENT=2
    LAYOUT_BLANK=12
    LAYOUT_SECTION=11
    LAYOUT_TWO_CONTENT=3
    MSO_TEXTBOX=17
    MSO_RECTANGLE=1
    MSO_ROUNDED_RECTANGLE=5
    MSO_OVAL=9
    MSO_RIGHT_ARROW=33
    MSO_LEFT_RIGHT_ARROW=37

    @property
    def slide_count(self):
        return int(self._run('return count of slides of ownedDoc'))

    def set_slide_size(self,width_pt=960,height_pt=540):
        self._mutable();_number(width_pt,positive=True);_number(height_pt,positive=True)
        if height_pt==540:
            body=f'set slide size of page setup of ownedDoc to slide size on screen\nset slide width of page setup of ownedDoc to {width_pt}'
        else:
            from .macos_powerpoint_script import compile_initial_slide_size
            body=compile_initial_slide_size(width_pt,height_pt)
        result=self._run(body+'\nreturn "SIZED"',mutation=True)
        if result=='EXISTING_SLIDES':
            raise PowerPointSessionCapabilityError(
                'Arbitrary initial slide size requires an empty presentation')

    def _role_font(self,target,role,size,color,default_size):
        _number(size,positive=True)
        family=None;preset=getattr(self,'_design_preset',None)
        if preset is not None:
            family,preset_size,preset_color=preset.get_font(role)
            if size==default_size:size=preset_size
            color=color or preset_color
        lines=[f'set font size of font of {target} to {_number(size,positive=True)}']
        if color is not None:lines.append(f'set font color of font of {target} to {_color(color)}')
        if family is not None:
            lines += [f'set font name of font of {target} to {apple_string(family)}',f'set east asian name of font of {target} to {apple_string(family)}']
        return lines

    def _semantic_slide(self,layout,lines):
        self._mutable()
        body=f'set currentSlide to make new slide at end of ownedDoc with properties {{layout:{_enum(layout,LAYOUT)}}}\n'+'\n'.join(lines)+'\nreturn slide index of currentSlide'
        index=int(self._run(body,mutation=True));self._current_slide=index
        return index

    def add_title_slide(self,title,subtitle="",title_size=40,sub_size=20,title_color=None):
        lines=[f'set content of text range of text frame of shape 1 of currentSlide to {apple_string(title)}']
        lines += self._role_font('text range of text frame of shape 1 of currentSlide','title',title_size,title_color,40)
        if subtitle:
            lines += [f'set content of text range of text frame of shape 2 of currentSlide to {apple_string(subtitle)}']
            lines += self._role_font('text range of text frame of shape 2 of currentSlide','subtitle',sub_size,None,20)
        return self._semantic_slide(1,lines)

    def add_section_slide(self,title):
        lines=[f'set content of text range of text frame of shape 1 of currentSlide to {apple_string(title)}']
        lines += self._role_font('text range of text frame of shape 1 of currentSlide','title',32,None,32)
        return self._semantic_slide(11,lines)

    def add_text_slide(self,title,body,title_size=32,body_size=18,title_color=None,body_color=None,bullets=True):
        if isinstance(body,list):
            if not bullets or any(not isinstance(item,str) for item in body):raise ValueError('List body requires text bullets')
            body='\r'.join(body)
        lines=[f'set content of text range of text frame of shape 1 of currentSlide to {apple_string(title)}',f'set content of text range of text frame of shape 2 of currentSlide to {apple_string(body)}']
        lines += self._role_font('text range of text frame of shape 1 of currentSlide','title',title_size,title_color,32)
        lines += self._role_font('text range of text frame of shape 2 of currentSlide','body',body_size,body_color,18)
        return self._semantic_slide(2,lines)

    def add_bullets_slide(self,title,items,title_size=32,body_size=18):
        return self.add_text_slide(title,items,title_size,body_size,bullets=True)

    def add_blank_slide(self):
        index=self._semantic_slide(12,[])
        return f'slide:{index}',index

    def set_background_color(self,slide_index,color):
        return self.apply_format_patch(f'slide:{_index(slide_index)}',follow_master_background=False,background={'visible':True,'color':color})

    def _make_shape(self,slide_index,native,left,top,width,height,lines,extra=''):
        self._mutable();index=_index(slide_index)
        bounds={k:_number(v,positive=k in {'width','height'}) for k,v in [('left',left),('top',top),('width',width),('height',height)]}
        name='wpscomposer-'+uuid4().hex
        body=f'set targetSlide to slide {index} of ownedDoc\nset currentShape to make new {native} at end of targetSlide with properties {{name:{apple_string(name)},left position:{bounds["left"]},top:{bounds["top"]},width:{bounds["width"]},height:{bounds["height"]}{extra}}}\n'+'\n'.join(lines)+'\nreturn "CREATED"'
        self._run(body,mutation=True)
        return f'slide:{index}/shape:@name={name}'

    def add_textbox(self,slide_index,text,left,top,width,height,size=18,bold=False,color=None,align=1,fill_color=None,shape_type=None):
        lines=[f'set content of text range of text frame of currentShape to {apple_string(text)}',f'set font size of font of text range of text frame of currentShape to {_number(size,positive=True)}',f'set bold of font of text range of text frame of currentShape to {_boolean(bold)}',f'set alignment of paragraph format of text range of text frame of currentShape to {_enum(align,ALIGN)}']
        if color is not None:lines.append(f'set font color of font of text range of text frame of currentShape to {_color(color)}')
        if fill_color is not None:lines += ['set visible of fill format of currentShape to true',f'set fore color of fill format of currentShape to {_color(fill_color)}']
        else:lines += ['set visible of fill format of currentShape to false','set transparency of line format of currentShape to 1']
        native='shape' if shape_type is not None else 'text box'
        extra=',auto shape type:'+_enum(shape_type,AUTOSHAPES) if shape_type is not None else ''
        return self._make_shape(slide_index,native,left,top,width,height,lines,extra)

    def add_shape(self,slide_index,shape_type,left,top,width,height,fill_color=None,line_color=None,text="",size=14):
        extra=',auto shape type:'+_enum(shape_type,AUTOSHAPES)
        lines=['set visible of fill format of currentShape to '+('true' if fill_color is not None else 'false')]
        if fill_color is not None:lines.append(f'set fore color of fill format of currentShape to {_color(fill_color)}')
        if line_color is None:lines.append('set transparency of line format of currentShape to 1')
        else:lines += ['set transparency of line format of currentShape to 0',f'set fore color of line format of currentShape to {_color(line_color)}']
        _number(size,positive=True);apple_string(text)
        if text:lines += [f'set content of text range of text frame of currentShape to {apple_string(text)}',f'set font size of font of text range of text frame of currentShape to {size}']
        return self._make_shape(slide_index,'shape',left,top,width,height,lines,extra)

    def add_image(self,slide_index,path,left,top,width=None,height=None):
        self._mutable();index=_index(slide_index);_number(left);_number(top)
        for v in (width,height):
            if v is not None:_number(v,positive=True)
        resource=Path(path).expanduser().resolve()
        if not resource.is_file() or resource.suffix.lower() not in {'.png','.jpg','.jpeg','.gif','.bmp','.tiff','.tif'}:raise ValueError('Invalid image resource')
        staged=self._job/('image-'+uuid4().hex+resource.suffix);copy_file_before_deadline(resource,staged,deadline=self._deadline)
        name='wpscomposer-'+uuid4().hex
        lines=[f'set currentPicture to make new picture at end of slide {index} of ownedDoc with properties {{name:{apple_string(name)},file name:{apple_string(str(staged))},link to file:false,save with document:true,left position:{left},top:{top}}}']
        if width is not None or height is not None:
            lines.append('set lock aspect ratio of currentPicture to '+('false' if width is not None and height is not None else 'true'))
            for key,value in [('width',width),('height',height)]:
                if value is not None:lines.append(f'set {key} of currentPicture to {value}')
        self._run('\n'.join(lines)+'\nreturn "IMAGE"',mutation=True)
        return f'slide:{index}/shape:@name={name}'

    def add_table(self,slide_index,rows,cols,left,top,width,height,data,header_shade="#4472C4",header_font="#FFFFFF",font_size=11):
        if type(rows)is not int or type(cols)is not int or not 1<=rows<=1000 or not 1<=cols<=1000 or rows*cols>10000:raise ValueError('Invalid table dimensions')
        _number(font_size,positive=True);_color(header_shade);_color(header_font)
        if not isinstance(data,(list,tuple)) or len(data)>rows or any(not isinstance(row,(list,tuple)) or len(row)>cols for row in data):raise ValueError('Table data exceeds dimensions')
        lines=[]
        for ri,row in enumerate(data,1):
            for ci,value in enumerate(row,1):
                lines += [f'set currentCell to get cell from table object of currentShape row {ri} column {ci}',f'set content of text range of text frame of shape of currentCell to {apple_string(str(value))}',f'set font size of font of text range of text frame of shape of currentCell to {font_size}']
                if ri==1:lines += ['set bold of font of text range of text frame of shape of currentCell to true',f'set font color of font of text range of text frame of shape of currentCell to {_color(header_font)}',f'set fore color of fill format of shape of currentCell to {_color(header_shade)}']
        return self._make_shape(slide_index,'shape table',left,top,width,height,lines,f',number of rows:{rows},number of columns:{cols}')

    def set_notes(self,slide_index,text):
        self._mutable();index=_index(slide_index);value=apple_string(text)
        self._run(f'set content of text range of text frame of shape 2 of notes page of slide {index} of ownedDoc to {value}\nreturn "NOTES"',mutation=True)

    def save_pptx(self,path):return self.save(path)

    def apply_design_preset(self,preset):
        from ..design_presets import DesignPreset
        from .._colors import resolve_color
        if not isinstance(preset,DesignPreset):raise TypeError('preset must be a DesignPreset instance')
        color=_color(resolve_color('bg',preset))
        for family,size,font_color in preset.fonts.values():apple_string(family);_number(size,positive=True);_color(font_color)
        self._mutable();self._run(f'set fore color of fill format of background of slide master of ownedDoc to {color}\nreturn "PRESET"',mutation=True)
        self._design_preset=preset

    def apply_layout_template(self,layout,preset=None):
        from ..layout_templates import LayoutTemplate,resolve_layout_colors
        if not isinstance(layout,LayoutTemplate):raise TypeError('layout must be a LayoutTemplate instance')
        if preset is not None:layout=resolve_layout_colors(layout,preset)
        self._mutable();lines=[];ignored=[]
        # Compile all static elements before changing the document.
        for elem in layout.elements:
            kind=elem.get('type')
            if kind not in {'bg','line','box','text'}:ignored.append(kind);continue
            if kind=='bg':
                lines += ['set follow master background of targetSlide to false','set visible of fill format of background of targetSlide to true',f'set fore color of fill format of background of targetSlide to {_color(elem.get("color","#FFFFFF"))}'];continue
            left=elem.get('x',0);top=elem.get('y',0);width=elem.get('w',100 if kind!='text' else 300);height=elem.get('h',3 if kind=='line' else 100 if kind=='box' else 50)
            for key,value in [('left',left),('top',top),('width',width),('height',height)]:_number(value,positive=key in {'width','height'})
            color=_color(elem.get('color','#EEEEEE' if kind=='box' else '#000000'))
            name=apple_string('wpscomposer-'+uuid4().hex)
            if kind=='line':
                midpoint=top+height/2
                lines += [f'set layoutShape to make new line shape at end of targetSlide with properties {{name:{name},begin line X:{left},begin line Y:{midpoint},end line X:{left+width},end line Y:{midpoint}}}',f'set line weight of line format of layoutShape to {max(height,1)}',f'set fore color of line format of layoutShape to {color}'];continue
            text=elem.get('text','');apple_string(text)
            if kind=='text' and not text:continue
            native='shape' if kind=='box' else 'text box';extra=',auto shape type:autoshape rectangle' if kind=='box' else ''
            lines += [f'set layoutShape to make new {native} at end of targetSlide with properties {{name:{name},left position:{left},top:{top},width:{width},height:{height}{extra}}}']
            if kind=='box':lines += [f'set fore color of fill format of layoutShape to {color}','set transparency of line format of layoutShape to 1']
            if text:
                size=12 if kind=='box' else elem.get('fs',18);_number(size,positive=True)
                lines += [f'set content of text range of text frame of layoutShape to {apple_string(text)}',f'set font size of font of text range of text frame of layoutShape to {size}']
                if kind=='text':
                    lines += [f'set font color of font of text range of text frame of layoutShape to {color}',f'set bold of font of text range of text frame of layoutShape to {_boolean(elem.get("bold",False))}']
                    if elem.get('align'):lines.append(f'set alignment of paragraph format of text range of text frame of layoutShape to {_enum(elem["align"],ALIGN)}')
        binding='set targetSlide to slide of view of document window 1 of ownedDoc'
        self._run(binding+'\n'+'\n'.join(lines)+'\nreturn "LAYOUT"',mutation=True)
        return {'ignored_compound_elements':ignored}

    def _validate_artifact(self,path,fmt):
        spec=ValidatorSpec.from_callable(validate_pdf) if fmt=='pdf' else ValidatorSpec.from_callable(validate_office_package,fmt)
        return validate_before_deadline(spec,path,self._deadline)

    def _destination(self,path,suffix):
        target=Path(path).expanduser().resolve()
        if target.suffix.lower()!=suffix:raise PowerPointSessionCapabilityError('Unexpected PowerPoint output extension')
        if target.exists():raise FileExistsError(target)
        return target

    def save(self,path):
        self._mutable()
        if self._attached:raise PowerPointSessionCapabilityError('Attached SaveAs would change document identity; use save_current')
        target=self._destination(path,'.pptx')
        self._run('save ownedDoc\nreturn "SAVED"',mutation=True)
        self._validate_artifact(Path(self._path),'pptx')
        published_state=snapshot_artifact_state(Path(self._path),deadline=self._deadline)
        publish_artifact(Path(self._path),target,overwrite=False,validator=lambda p:self._validate_artifact(p,'pptx'),deadline=self._deadline)
        current_state=snapshot_artifact_state(target,deadline=self._deadline)
        if current_state.sha256!=published_state.sha256:raise RuntimeError('Logical PowerPoint destination changed after publication')
        self._logical=target;self._logical_state=current_state
        return str(target)

    def preflight_save(self,output=None,*,overwrite=False):
        self._mutable()
        if output is not None:
            self._destination(output,'.pptx')
            if self._attached:raise PowerPointSessionCapabilityError('Native attached save-copy preserving document identity is unavailable')
            return
        if not Path(str(self._path)).is_absolute():raise PowerPointSessionCapabilityError('Unsaved attached presentation cannot be saved without an explicit supported destination')
        if self._native_read_only:raise PermissionError('Bound native presentation is read-only')
        if self._attached and (not Path(str(self._path)).is_file() or not os.access(self._path,os.W_OK)):
            raise PowerPointSessionCapabilityError('Bound saved presentation path is unavailable or not writable')
        if self._attached:validate_native_input(Path(self._path),'presentation',deadline=self._deadline)
        if self._attached:
            self._verify_source_digest()
        else:
            if self._logical is None:raise ValueError('New PowerPoint session requires an explicit save destination')
            if self._logical_state is None:
                # A source digest may predate logical tracking on an already
                # initialized session; validate it before capturing state.
                self._verify_source_digest()
                self._logical_state=snapshot_artifact_state(self._logical,deadline=self._deadline)
            try:current=snapshot_artifact_state(self._logical,deadline=self._deadline)
            except (OSError,ValueError):raise ValueError('Logical PowerPoint destination changed during session') from None
            if current!=self._logical_state:raise ValueError('Logical PowerPoint destination changed during session')

    def save_current(self):
        self.preflight_save()
        if not Path(str(self._path)).is_absolute():raise PowerPointSessionCapabilityError('Saving an unsaved attached presentation requires explicit save destination support')
        self._run('save ownedDoc\nreturn "SAVED"',mutation=True)
        if self._attached:
            self._bound_digest=self._digest(self._path)
            return str(self._path)
        self._validate_artifact(Path(self._path),'pptx')
        published_state=snapshot_artifact_state(Path(self._path),deadline=self._deadline)
        publish_artifact(Path(self._path),self._logical,overwrite=True,
                         validator=lambda p:self._validate_artifact(p,'pptx'),
                         deadline=self._deadline,expected_destination=self._logical_state)
        current_state=snapshot_artifact_state(self._logical,deadline=self._deadline)
        if current_state.sha256!=published_state.sha256:raise RuntimeError('Logical PowerPoint destination changed after publication')
        self._logical_state=current_state
        return str(self._logical)

    def save_copy(self,path):
        if self._attached:raise PowerPointSessionCapabilityError('Native attached save-copy preserving path and saved-state not verified')
        logical,state=self._logical,self._logical_state
        try:return self.save(path)
        finally:self._logical,self._logical_state=logical,state

    def export_pdf(self,path):
        if self._attached:raise PowerPointSessionCapabilityError('Attached PDF export preserving unsaved state is unverified')
        target=self._destination(path,'.pdf')
        native=self._job/('export-'+uuid4().hex+'.pdf')
        prefix=f'set pdfHFS to (POSIX file {apple_string(str(native))}) as text\n'
        self._run('save ownedDoc in pdfHFS as save as PDF\nreturn "PDF"',prefix=prefix,mutation=True)
        return str(publish_artifact(native,target,overwrite=False,validator=lambda p:self._validate_artifact(p,'pdf'),deadline=self._deadline))

    def is_bound_to(self,path):
        if not self._entered:return False
        actual=self._run('return full name of ownedDoc')
        return actual==str(Path(path).expanduser().resolve())

    def _release(self):
        if self._lock:self._lock.close();self._lock=None

    def close(self,save_changes=False):
        if save_changes:
            self._mutable()
            if self._entered and not self._attached and not self._uncertain:self.save_current()
        try:
            if self._entered and not self._attached and not self._uncertain:
                self._run('close ownedDoc saving no\nreturn "CLOSED"',mutation=True)
        finally:
            self._entered=False;self._release()
