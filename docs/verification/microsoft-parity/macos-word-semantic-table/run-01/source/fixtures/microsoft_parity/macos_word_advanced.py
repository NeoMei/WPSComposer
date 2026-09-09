"""Opt-in native Mac Word advanced parity acceptance; no OOXML writing.

Run with --output-dir pointing at a new evidence directory. All native document
operations share MacWordAdapter's job lock, private staging and total deadline.
The synthetic unsaved sentinel is closed only when its full snapshot is intact.
"""
from __future__ import annotations

import argparse
from dataclasses import replace
import hashlib
import json
import math
from pathlib import Path
import shutil
import subprocess
import time
from uuid import uuid4
import zipfile
from xml.etree import ElementTree as ET

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, apple_string, parse_result


def build_case(output):
    from PIL import Image
    for name, color in [('left.png','#115599'), ('right.png','#FFBB33')]:
        Image.new('RGB', (240, 120), color).save(output/name)
    fallback = '[REFERENCE_UNRESOLVED 引用目标未解析]'
    markdown = ('# Native Word advanced acceptance\n\n## Native equation\n\n'
                '$$\nx^2+\\frac{a}{b}\n$$\n\nAFTER-EQUATION\n\n'
                '## Landscape merge\n\n| A | B | C |\n|---|---|---|\n| MERGE | | OUTSIDE |\n| '+fallback+' | middle | right |\n\n'
                'AFTER-LANDSCAPE\n\n:::figure {caption="Two native images" layout="columns"}\n![Left](left.png)\n![Right](right.png)\n:::\n\n'
                'COLORED-STYLE\n\nBOLD-SEGMENT / STRIKE-SEGMENT\n\n- Custom bullet one\n- Custom bullet two\n\nEND-OF-DOCUMENT\n')
    build = build_longform_generation(markdown, base_dir=str(output))
    ops = []
    for operation in build.plan.operations:
        a = dict(operation.args)
        if operation.op == 'writer.add_semantic_table':
            a.update(orientation='landscape', includePreviousHeading=True, continuousExit=True,
                     merges=[dict(top=2,left=1,bottom=2,right=2)],
                     cellDegradations=[dict(row=3,column=1,code='REFERENCE_UNRESOLVED',fallbackText=fallback)])
        elif operation.op == 'writer.ensure_styles':
            a['styles'] = [dict(s) for s in a['styles']] + [dict(name='TaskColored', type='paragraph', basedOn='Body Text', color='#115599', shading='#E6F0FA', underline=True, strikethrough=True, leftIndent=18, rightIndent=9, lineSpacing=20, lineSpacingRule='exactly',leftBorder=True,borderColor='#115599')]
        elif operation.op == 'writer.add_paragraph' and a['text'] == 'COLORED-STYLE':
            a['style'] = 'TaskColored'
        elif operation.op == 'writer.add_paragraph' and a['text'] == 'BOLD-SEGMENT / STRIKE-SEGMENT':
            a['spans'] = [dict(text='BOLD-SEGMENT',bold=True),dict(text=' / '),dict(text='STRIKE-SEGMENT',strikethrough=True)]
        elif operation.op == 'writer.configure_section' and a['role'] == 'body':
            a.update(pageSize='A4', margins=dict(top=72,bottom=72,left=72,right=72), footerText='CONFIDENTIAL ', pageNumberFormat='arabic', linkToPreviousFooter=False)
        ops.append(replace(operation,args=a))
    return replace(build, plan=replace(build.plan,operations=tuple(ops)))


def inspect_package(path):
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'm':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
    with zipfile.ZipFile(path) as package:
        doc = ET.fromstring(package.read('word/document.xml'))
        styles = ET.fromstring(package.read('word/styles.xml'))
        text = ''.join(doc.itertext())
        formulas = doc.findall('.//m:oMath',ns)
        custom = next((s for s in styles.findall('w:style',ns) if s.find('w:name',ns) is not None and s.find('w:name',ns).get('{'+ns['w']+'}val') == 'TaskColored'),None)
        footer_text = ''.join(package.read(n).decode() for n in package.namelist() if n.startswith('word/footer'))
        formula_columns=doc.find('.//w:tbl/w:tblGrid',ns)
        widths=[int(x.get('{'+ns['w']+'}w')) for x in formula_columns] if formula_columns is not None else []
        return dict(formula_middle_has_body_width=len(widths)==3 and widths[1]>widths[0]*3 and widths[1]>widths[2]*3,
                    native_math=len(formulas)==1, native_superscript=doc.find('.//m:sSup',ns) is not None,
                    native_fraction=doc.find('.//m:f',ns) is not None,
                    math_did_not_absorb_following_text=all('AFTER-EQUATION' not in ''.join(x.itertext()) for x in formulas),
                    merged_grid_span=doc.find('.//w:gridSpan',ns) is not None,
                    merge_outside_retained='OUTSIDE' in text,
                    landscape=any(x.get('{'+ns['w']+'}orient')=='landscape' for x in doc.findall('.//w:pgSz',ns)),
                    portrait_restored=doc.findall('.//w:pgSz',ns)[-1].get('{'+ns['w']+'}orient')!='landscape',
                    pictures_in_two_real_cells=any(len(t.findall('w:tr/w:tc',ns))==2 and all(len(c.findall('.//w:drawing',ns))==1 for c in t.findall('w:tr/w:tc',ns)) for t in doc.findall('.//w:tbl',ns)),
                    embedded_images=len([n for n in package.namelist() if n.startswith('word/media/')])==2,
                    colored_style=custom is not None and custom.find('.//w:color',ns).get('{'+ns['w']+'}val')=='115599',
                    document_tail='END-OF-DOCUMENT' in text,
                    footer_text_and_page='CONFIDENTIAL' in footer_text and 'PAGE' in footer_text)


def run(output, timeout=180, force_math_failure=False):
    if isinstance(timeout,bool) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('timeout must be finite and positive')
    source_digest = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    compiler_digest = hashlib.sha256(Path('skills/WPSComposer/scripts/msoffice/macos_script.py').read_bytes()).hexdigest()
    if output.exists():
        raise FileExistsError(output)
    output.mkdir(parents=True)
    build = build_case(output)
    if force_math_failure:
        from skills.WPSComposer.scripts.longform.native_math import convert_restricted_latex
        operations=[]
        for op in build.plan.operations:
            if op.op == 'writer.add_equation':
                descriptor=convert_restricted_latex(op.args['source'])
                args={k:v for k,v in op.args.items() if k!='source'}
                args.update(renderMode='native-m4',content={'nativeMath':{'syntax':descriptor.syntax,'linearText':descriptor.linear_text,'sourceHash':descriptor.source_hash}})
                op=replace(op,args=args,failure_policy={'mode':'degrade','recoverableCodes':['EQUATION_INSERT_FAILED'],'fallback':'explicit-image-then-source-notice'})
            operations.append(op)
        build=replace(build,plan=replace(build.plan,operations=tuple(operations)))
    adapter = MacWordAdapter(build)
    deadline = time.monotonic()+timeout
    report = {'passed':False, 'checks':{}, 'issues':[], 'runtime_staging':None}
    sentinel = None
    try:
        adapter._ensure_started(deadline)
        report['runtime_staging'] = str(adapter.staging_root)
        marker = 'WPSC-ADVANCED-SENTINEL-' + uuid4().hex
        raw = adapter._run('''with timeout of 15 seconds
 tell application "Microsoft Word"
  set sentinelDoc to make new document
  set initialText to content of text object of sentinelDoc
  set initialName to name of sentinelDoc
  set initialPath to posix full name of sentinelDoc
  set initialSaved to saved of sentinelDoc
  try
   set content of text object of sentinelDoc to '''+apple_string(marker)+'''
   set sentinelText to content of text object of sentinelDoc
   return name of sentinelDoc & tab & posix full name of sentinelDoc & tab & (saved of sentinelDoc as text)
  on error errText number errNumber
   if name of sentinelDoc is initialName and posix full name of sentinelDoc is initialPath and saved of sentinelDoc is initialSaved and content of text object of sentinelDoc is initialText then close sentinelDoc saving no
   error errText number errNumber
  end try
 end tell
end timeout''',deadline)
        parts = raw.strip().split('\t')
        sentinel = (parts[-3], parts[-2], parts[-1], marker)
        # Resource payloads stay within Word's private container staging.
        from skills.WPSComposer.scripts.longform.pipeline import _build_executor_resources
        resources = {}
        for resource in _build_executor_resources(build.base_dir,build.preflight):
            path = adapter.staging_root/(hashlib.sha256(resource.id.encode()).hexdigest()[:16]+'.png')
            path.write_bytes(resource.payload_bytes)
            resources[resource.id]=path
        target = adapter.staging_root/'advanced.docx'
        compiled = compile_plan(build.plan,resources,target,timeout=max(1,deadline-time.monotonic()-2))
        if force_math_failure:
            compiled=replace(compiled,source=compiled.source.replace('build up ownMath','error "Task-owned injected native math failure"'))
        report['compiled_sha256']=hashlib.sha256(compiled.source.encode()).hexdigest()
        report['injected_math_failure']=force_math_failure
        traced = []
        picture_trace = False
        for step, line in enumerate(compiled.source.splitlines(), 1):
            if line.startswith('log "WPSC_OP:'):
                picture_trace = 'add_captioned_figure' in line
            if picture_trace:
                traced.append(f'log "WPSC_PICTURE_STEP:{step}"')
            traced.append(line)
        compiled = replace(compiled, source='\n'.join(traced))
        (output/'compiled.applescript').write_text(compiled.source)
        result = subprocess.run(['/usr/bin/osacompile','-o',str(output/'compiled.scpt'),str(output/'compiled.applescript')],capture_output=True,text=True,timeout=20)
        (output/'compile.log').write_text(result.stdout+'\n'+result.stderr)
        if result.returncode:
            raise RuntimeError('AppleScript compilation failed')
        raw = adapter._run(compiled.source,deadline)
        outcome = parse_result(raw,compiled.nodes,target,compiled.operations,compiled.issues)
        (output/'generation.log').write_text(raw)
        shutil.copyfile(target,output/'advanced.docx')
        pdf = adapter.export_pdf(target,deadline)
        shutil.copyfile(pdf,output/'advanced.pdf')
        report['checks'] = inspect_package(output/'advanced.docx')
        report['pagination_nodes'] = len(outcome.pagination_map.nodes)
        report['issues'] = [issue.to_dict() for issue in outcome.issues]
        expected_issues={'REFERENCE_UNRESOLVED'} | ({'EQUATION_INSERT_FAILED'} if force_math_failure else set())
        report['checks']['no_unexpected_issues'] = {issue.code for issue in outcome.issues} == expected_issues
        if force_math_failure:
            with zipfile.ZipFile(output/'advanced.docx') as package:
                content=ET.fromstring(package.read('word/document.xml'))
                ns={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
                report['checks']['source_fallback_retained']=not content.findall('.//m:oMath',ns) and 'EQUATION_INSERT_FAILED' in ''.join(content.itertext())
            for key in ('native_math','native_superscript','native_fraction'):
                report['checks'].pop(key)
        import fitz
        with fitz.open(output/'advanced.pdf') as pdf_doc:
            report['pdf_pages'] = len(pdf_doc)
            pdf_text = ''.join(page.get_text() for page in pdf_doc)
            report['checks']['pdf_tail_retained'] = 'END-OF-DOCUMENT' in pdf_text
            for i,page in enumerate(pdf_doc):
                page.get_pixmap(matrix=fitz.Matrix(1,1)).save(output/f'page-{i+1}.png')
    except Exception as error:
        report['error'] = str(error)
    finally:
        if sentinel is not None:
            name,path,saved,marker=sentinel
            cleanup = f'''with timeout of 15 seconds
 tell application "Microsoft Word"
  set sentinelDoc to document {apple_string(name)}
  if posix full name of sentinelDoc is not {apple_string(path)} then error "Sentinel path changed"
  if (saved of sentinelDoc as text) is not {apple_string(saved)} then error "Sentinel saved state changed"
  if content of text object of sentinelDoc is not {apple_string(marker)} & return then error "Sentinel text changed"
  close sentinelDoc saving no
 end tell
end timeout
return "SENTINEL_PRESERVED"'''
            try:
                # Cleanup is independent and bounded even when job deadline expired.
                result = subprocess.run(['/usr/bin/osascript','-e',cleanup],capture_output=True,text=True,timeout=20)
                (output/'sentinel.log').write_text(result.stdout+'\n'+result.stderr)
                report['checks']['unsaved_sentinel_preserved'] = result.returncode == 0 and result.stdout.strip()=='SENTINEL_PRESERVED'
            except Exception as error:
                report['sentinel_cleanup_error']=str(error)
        if adapter.staging_root:
            shutil.copytree(adapter.staging_root,output/'runtime',dirs_exist_ok=True)
        adapter.close()
        report['passed'] = not report.get('error') and bool(report['checks']) and all(report['checks'].values())
        report['source_sha256'] = source_digest
        report['compiler_sha256'] = compiler_digest
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps(report,ensure_ascii=False,indent=2))
    return 0 if report['passed'] else 1


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir',type=Path,required=True)
    parser.add_argument('--timeout',type=float,default=180)
    parser.add_argument('--force-math-failure',action='store_true',help='Inject failure only in this task-owned fixture equation, verifying native source recovery')
    args=parser.parse_args()
    raise SystemExit(run(args.output_dir.resolve(),args.timeout,args.force_math_failure))
