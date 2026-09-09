"""Privacy-safe pagination facts from the exactly bound native Word document.

Positions are Word offsets (UTF-16 units), never Python string offsets. Ordinary
Start/End objects retain the frozen direct API's coercion semantics; existing
session-tagged degradation handles additionally retain their ownership contract.
"""
from __future__ import annotations

import math

from ..longform.privacy import redact_private_text
from ..writer import NativeWriterObjectError
from .macos_script import apple_string
from .macos_word_degradation import NativeDegradationRange


_VISUAL_OPS = {
    'writer.add_heading', 'writer.add_captioned_figure',
    'writer.add_semantic_table', 'writer.add_equation',
    'writer.add_degradation_notice', 'writer.add_document_quality_notice',
}


def _failed():
    raise NativeWriterObjectError(
        'PAGINATION_SNAPSHOT_FAILED', 'pagination snapshot failed'
    ) from None


def pagination_fragment_for_bookmark(session, node_id, bookmark_name):
    """Return frozen M5 first-paragraph point geometry without editing text."""
    try:
        if isinstance(bookmark_name, (int, float)):
            if not math.isfinite(bookmark_name):
                raise ValueError("invalid bookmark ordinal")
            name = str(int(bookmark_name)) if isinstance(bookmark_name, int) else repr(bookmark_name)
        else:
            name = apple_string(bookmark_name)
        rows = session._execute([
            f'set paginationBookmark to bookmark {name} of boundDoc',
            'set paginationRange to text object of paragraph 1 of text object of paginationBookmark',
            'set paginationPage to get range information paginationRange information type active end page number',
            'set paginationX to get range information paginationRange information type horizontal position relative to page',
            'set paginationY to get range information paginationRange information type vertical position relative to page',
            'set nativeRows to {{"bookmark",start of content of paginationRange,'
            'end of content of paginationRange,paginationPage,paginationX,paginationY}}',
        ])
        if (not isinstance(rows, list) or len(rows) != 1
                or not isinstance(rows[0], list) or len(rows[0]) != 6
                or rows[0][0] != 'bookmark'):
            raise ValueError('invalid pagination acknowledgement')
        _, start, end, page, x, y = rows[0]
        start, end, page = int(start), int(end), int(page)
        x, y = float(x), float(y)
        fragment = {'page': page}
        if all(math.isfinite(value) and value >= 0 for value in (x, y)):
            fragment['bounds'] = [x, y, x + 1.0, y + 12.0]
        return {
            'nodeId': redact_private_text(str(node_id)),
            'story': 'main', 'sections': ['body'],
            'pageStart': page, 'pageEnd': page,
            'range': f'{start}:{end}', 'fragments': [fragment],
        }
    except Exception:
        _failed()


def _tracked_ranges(session, tracked_ranges):
    prepared, seen = [], set()
    for tracked in tracked_ranges:
        node_id = str(tracked.get('nodeId') or '')
        if not node_id or node_id in seen:
            continue
        rng = tracked['range']
        if isinstance(rng, NativeDegradationRange):
            if rng.session_id != getattr(session, '_degradation_session_id', None):
                raise ValueError('foreign native range')
        start, end = int(rng.Start), int(rng.End)
        if start < 0 or end < start:
            raise ValueError('invalid native range')
        prepared.append((node_id, start, end,
                         str(tracked.get('role') or 'body'),
                         tracked.get('op') in _VISUAL_OPS))
        seen.add(node_id)
    return prepared


def _map_commands(prepared):
    lines = ['repaginate boundDoc', 'set paginationSetup to page setup of boundDoc']
    properties = (
        ('paginationWidth', 'page width', 595.28),
        ('paginationHeight', 'page height', 841.89),
        ('paginationLeft', 'left margin', 72.0),
        ('paginationRight', 'right margin', 72.0),
        ('paginationTop', 'top margin', 72.0),
        ('paginationBottom', 'bottom margin', 72.0),
    )
    for variable, prop, default in properties:
        # Match getattr's missing-property default, never hide transport errors.
        lines += [f'set {variable} to {default}', 'try',
                  f'set {variable} to {prop} of paginationSetup',
                  'on error paginationError number paginationNumber',
                  'if paginationNumber is not -1728 then error paginationError number paginationNumber',
                  'end try']
    lines += [
        'set nativeRows to {{"setup",paginationWidth,paginationHeight,paginationLeft,'
        'paginationRight,paginationTop,paginationBottom}}',
        'set paginationContentEnd to (end of content of text object of boundDoc) - 1',
        'if paginationContentEnd < 0 then set paginationContentEnd to 0',
    ]
    for index, (_, start, end, _, _) in enumerate(prepared):
        last_position = end - 1 if end > start else end
        lines += [
            f'set paginationFirst to {start}',
            'if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd',
            f'set paginationLast to {max(0, last_position)}',
            'if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd',
            'set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst',
            'set paginationLastRange to create range boundDoc start paginationLast end paginationLast',
            'set paginationFirstPage to get range information paginationFirstRange information type active end page number',
            'set paginationLastPage to get range information paginationLastRange information type active end page number',
            'set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page',
            'set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page',
            'set paginationLastY to get range information paginationLastRange information type vertical position relative to page',
            f'set end of nativeRows to {{"range",{index},paginationFirstPage,paginationLastPage,'
            'paginationFirstX,paginationFirstY,paginationLastY}',
        ]
    return lines


def _map_snapshot(prepared, rows):
    if (not isinstance(rows, list) or len(rows) != len(prepared) + 1
            or not isinstance(rows[0], list) or len(rows[0]) != 7
            or rows[0][0] != 'setup'):
        raise ValueError('invalid pagination acknowledgement')
    page_width, page_height = (float(v) for v in rows[0][1:3])
    left_margin, right_margin, top_margin, bottom_margin = (
        max(0.0, float(v)) for v in rows[0][3:7]
    )
    nodes = []
    for index, (node_id, start, end, role, visual) in enumerate(prepared):
        row = rows[index + 1]
        if (not isinstance(row, list) or len(row) != 7
                or row[0] != 'range' or type(row[1]) is not int or row[1] != index):
            raise ValueError('invalid pagination acknowledgement')
        first_page, last_page = int(row[2]), int(row[3])
        if first_page < 1 or last_page < first_page:
            raise ValueError('invalid native page span')
        first_x, first_y, last_y = (float(v) for v in row[4:7])
        usable_page = (
            100.0 <= page_width <= 2000.0
            and 100.0 <= page_height <= 2000.0
            and left_margin + right_margin < page_width
            and top_margin + bottom_margin < page_height
        )
        usable_points = (
            usable_page
            and all(math.isfinite(v) for v in (first_x, first_y, last_y))
            and 0.0 <= first_x <= page_width
            and 0.0 <= first_y <= page_height
            and 0.0 <= last_y <= page_height
        )
        fragments = []
        for page in range(first_page, last_page + 1):
            fragment = {'page': page}
            if visual and usable_points:
                x0 = max(0.0, first_x) if page == first_page and math.isfinite(first_x) else left_margin
                y0 = max(0.0, first_y) if page == first_page and math.isfinite(first_y) else top_margin
                x1 = max(x0 + 1.0, page_width - right_margin)
                y1 = (
                    max(y0 + 1.0, min(page_height, last_y + 12.0))
                    if page == last_page and math.isfinite(last_y)
                    else max(y0 + 1.0, page_height - bottom_margin)
                )
                if not all(math.isfinite(v) for v in (x0, y0, x1, y1)):
                    raise ValueError('invalid native bounds')
                fragment['bounds'] = [x0, y0, x1, y1]
            fragments.append(fragment)
        nodes.append({
            'nodeId': redact_private_text(node_id),
            'story': 'main', 'sections': [role],
            'pageStart': first_page, 'pageEnd': last_page,
            'range': f'{start}:{end}', 'fragments': fragments,
        })
    return {'version': 'M5-v1', 'nodes': nodes}


def pagination_map_for_ranges(session, tracked_ranges):
    """Repaginate and return M5 page spans without returning document text."""
    try:
        prepared = _tracked_ranges(session, tracked_ranges)
        return _map_snapshot(prepared, session._execute(_map_commands(prepared)))
    except Exception:
        _failed()
