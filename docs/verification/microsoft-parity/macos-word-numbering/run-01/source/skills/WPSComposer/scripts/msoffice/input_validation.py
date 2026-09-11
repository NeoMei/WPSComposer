"""Read-only OPC input gate shared by native Word, Excel and PowerPoint hosts.

This is a conservative macro-free OOXML gate, not an Office sandbox. It rejects
active declarations and embedded/linked resources before native opening, while
allowing ordinary document hyperlinks. It does not inspect legacy binary Office
formats. XML parsing is bounded; no package contents are extracted or executed.
"""
from __future__ import annotations

import math
from pathlib import Path
import posixpath
import re
import time
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

PUBLIC_ERROR = 'Unsupported or unsafe native Office input.'
TIMEOUT_ERROR = 'Native Office input validation timed out.'
MAX_XML_PART_BYTES = 16 * 1024 * 1024
MAX_TOTAL_XML_BYTES = 64 * 1024 * 1024
MAX_MEMBERS = 16384
_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
_OFFICE_BASES = ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
                 'http://purl.oclc.org/ooxml/officeDocument/relationships/')
_MAIN = {
    'writer': ('.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
    'spreadsheet': ('.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
    'presentation': ('.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
}
_ACTIVE_TOKENS = ('macroenabled', 'vbaproject', 'vbadata', 'macrosheet', 'activex',
                  'oleobject', 'externallink', 'connections', 'querytable')
_ACTIVE_REL_ENDINGS = {'package', 'control', 'ctrlprop', 'attachedtemplate', 'afchunk'}


def _reject():
    raise ValueError(PUBLIC_ERROR)


def _check_deadline(deadline):
    if deadline is None:
        return
    if isinstance(deadline, bool) or not isinstance(deadline, (int, float)) or not math.isfinite(deadline):
        _reject()
    if time.monotonic() >= deadline:
        raise TimeoutError(TIMEOUT_ERROR)


def _part_name(name):
    decoded = unquote(name)
    if (not decoded or '\\' in decoded or decoded.startswith('/') or
            any(ord(char) < 32 for char in decoded) or
            any(piece in ('', '.', '..') for piece in decoded.split('/')) or
            urlsplit(decoded).scheme or '?' in decoded or '#' in decoded):
        _reject()
    return decoded


def _active(value):
    folded = value.strip().lower()
    return any(token in folded for token in _ACTIVE_TOKENS)


def _read_xml(archive, info, budget, deadline):
    _check_deadline(deadline)
    if info.file_size > MAX_XML_PART_BYTES or budget[0] + info.file_size > MAX_TOTAL_XML_BYTES:
        _reject()
    budget[0] += info.file_size
    with archive.open(info) as stream:
        data = stream.read(MAX_XML_PART_BYTES + 1)
    if len(data) > MAX_XML_PART_BYTES or len(data) != info.file_size:
        _reject()
    # Null stripping also recognizes these declarations in UTF-16/UTF-32 XML.
    declaration_check = data.replace(b'\x00', b'').upper()
    if b'<!DOCTYPE' in declaration_check or b'<!ENTITY' in declaration_check:
        _reject()
    _check_deadline(deadline)
    root = ET.fromstring(data)
    _check_deadline(deadline)
    return root


def _resolve_target(rels_name, target):
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or not parsed.path or '\\' in target:
        _reject()
    decoded = unquote(parsed.path)
    if rels_name == '_rels/.rels':
        base = ''
    else:
        directory, filename = posixpath.split(rels_name)
        if not directory.endswith('_rels') or not filename.endswith('.rels'):
            _reject()
        base = posixpath.dirname(directory)
    resolved = posixpath.normpath(decoded.lstrip('/') if decoded.startswith('/') else posixpath.join(base, decoded))
    return _part_name(resolved)


def validate_spreadsheet_formula(formula):
    """Reject known automatically executing external/automation formulas.

    This check does not execute or evaluate expressions. Quoted text literals
    are excluded so a label containing a function name remains ordinary text.
    """
    code = re.sub(r'"(?:[^"]|"")*"', '""', str(formula))
    functions = re.findall(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', code)
    blocked = {'RTD', 'WEBSERVICE', 'IMAGE', 'CALL', 'REGISTER', 'REGISTER.ID', 'EXEC',
               'RUN', 'EVALUATE', 'DDE', 'SQL.REQUEST'}
    for function in functions:
        name = function.upper()
        while name.startswith(('_XLFN.', '_XLWS.')):
            name = name.split('.', 1)[1]
        if name in blocked:
            _reject()
    if '|' in code:
        _reject()


def _validate_spreadsheet_fields(root):
    namespaces = ('http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'http://purl.oclc.org/ooxml/spreadsheetml/main',
                  'http://schemas.microsoft.com/office/excel/2006/main')
    tags = {'{' + ns + '}' + name for ns in namespaces
            for name in ('f', 'formula', 'formula1', 'formula2', 'definedName')}
    for element in root.iter():
        if element.tag in tags:
            validate_spreadsheet_formula(''.join(element.itertext()))


def _validate_word_fields(root):
    # Word can execute fields without a VBA part or an external OPC relation.
    namespaces = ('http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                  'http://purl.oclc.org/ooxml/wordprocessingml/main')
    blocked = {'DDE', 'DDEAUTO', 'INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'DATABASE', 'MACROBUTTON'}

    def check(chunks):
        instruction = ''.join(chunks).strip().split(None, 1)
        if instruction and instruction[0].upper() in blocked:
            _reject()

    for ns in namespaces:
        tag = '{' + ns + '}'
        stack, loose = [], []
        for element in root.iter():
            if element.tag == tag + 'p':
                check(loose)
                loose = []
            elif element.tag == tag + 'fldSimple':
                check([element.get(tag + 'instr', '')])
            elif element.tag == tag + 'fldChar':
                kind = element.get(tag + 'fldCharType')
                if kind == 'begin':
                    check(loose)
                    loose = []
                    stack.append([])
                elif kind in ('separate', 'end') and stack:
                    check(stack[-1])
                    if kind == 'end':
                        stack.pop()
            elif element.tag == tag + 'instrText':
                (stack[-1] if stack else loose).append(element.text or '')
        check(loose)
        for chunks in stack:
            check(chunks)


def _validate_package(path, component, deadline):
    if component not in _MAIN or path.suffix.lower() != _MAIN[component][0]:
        _reject()
    budget = [0]
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        if len(infos) > MAX_MEMBERS:
            _reject()
        parts = {}
        folded_names = set()
        for info in infos:
            _check_deadline(deadline)
            if info.is_dir():
                _part_name(info.filename.rstrip('/'))
                continue
            name = _part_name(info.filename)
            if name.casefold() in folded_names or info.flag_bits & 1:
                _reject()
            folded_names.add(name.casefold())
            parts[name] = info
        if '[Content_Types].xml' not in parts or '_rels/.rels' not in parts:
            _reject()
        types = _read_xml(archive, parts['[Content_Types].xml'], budget, deadline)
        if types.tag != '{' + _CT + '}Types':
            _reject()
        defaults, overrides = {}, {}
        for entry in types:
            _check_deadline(deadline)
            content_type = entry.get('ContentType', '').strip()
            if not content_type or _active(content_type):
                _reject()
            if entry.tag == '{' + _CT + '}Default':
                extension = entry.get('Extension', '').lower()
                if not extension or extension in defaults or any(c in extension for c in '/\\.'):
                    _reject()
                defaults[extension] = content_type
            elif entry.tag == '{' + _CT + '}Override':
                raw_name = entry.get('PartName', '')
                if not raw_name.startswith('/'):
                    _reject()
                name = _part_name(raw_name[1:])
                if name in overrides or name not in parts:
                    _reject()
                overrides[name] = content_type
            else:
                _reject()
        content_types = {}
        for name in parts:
            if name == '[Content_Types].xml':
                continue
            content_type = overrides.get(name, defaults.get(name.rsplit('.', 1)[-1].lower()))
            if content_type is None or _active(content_type):
                _reject()
            content_types[name] = content_type
        main_targets = []
        for name, info in parts.items():
            if name == '[Content_Types].xml':
                continue
            content_type = content_types[name].lower()
            is_relationship = content_type == 'application/vnd.openxmlformats-package.relationships+xml'
            if name.lower().endswith('.rels') and not is_relationship:
                _reject()
            is_xml = name.lower().endswith(('.xml', '.rels')) or content_type.endswith(('+xml', '/xml'))
            if not is_xml:
                continue
            root = _read_xml(archive, info, budget, deadline)
            if not is_relationship:
                if component == 'writer':
                    _validate_word_fields(root)
                elif component == 'spreadsheet':
                    _validate_spreadsheet_fields(root)
                continue
            if root.tag != '{' + _REL + '}Relationships':
                _reject()
            relationship_ids = set()
            for entry in root:
                _check_deadline(deadline)
                relation_id = entry.get('Id')
                rel_type = entry.get('Type', '').strip()
                target = entry.get('Target', '')
                if (entry.tag != '{' + _REL + '}Relationship' or not relation_id or
                        relation_id in relationship_ids or not rel_type or not target):
                    _reject()
                relationship_ids.add(relation_id)
                if _active(rel_type) or rel_type.rsplit('/', 1)[-1].lower() in _ACTIVE_REL_ENDINGS:
                    _reject()
                mode = entry.get('TargetMode', 'Internal')
                if mode == 'External':
                    if rel_type not in tuple(base + 'hyperlink' for base in _OFFICE_BASES):
                        _reject()
                    continue
                if mode != 'Internal':
                    _reject()
                resolved = _resolve_target(name, target)
                if resolved not in parts:
                    _reject()
                if name == '_rels/.rels' and rel_type in tuple(base + 'officeDocument' for base in _OFFICE_BASES):
                    main_targets.append(resolved)
        if len(main_targets) != 1 or content_types[main_targets[0]].lower() != _MAIN[component][1]:
            _reject()


def validate_native_input(path, component, *, deadline=None) -> None:
    """Validate DOCX/XLSX/PPTX before native open; invalid or active input raises ValueError.

    Public exception messages never contain names, XML, URLs or document content.
    The deadline is monotonic; timeout raises TimeoutError. Parser limits are
    16 MiB per XML part, 64 MiB aggregate XML and 16,384 package members.
    """
    _check_deadline(deadline)
    try:
        _validate_package(Path(path), component, deadline)
        _check_deadline(deadline)
    except TimeoutError:
        raise TimeoutError(TIMEOUT_ERROR) from None
    except (ValueError, TypeError, OSError, KeyError, ET.ParseError, zipfile.BadZipFile, RuntimeError):
        raise ValueError(PUBLIC_ERROR) from None
