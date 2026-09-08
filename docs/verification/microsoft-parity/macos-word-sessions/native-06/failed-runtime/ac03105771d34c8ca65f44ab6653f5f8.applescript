use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex
with timeout of 60 seconds
tell application "/Applications/Microsoft Word.app"
set nativeRows to {}
set boundDoc to document "document.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xgjv080d/document.docx" then error "WPSC_STALE_DOCUMENT"
set end of nativeRows to {"doc", 0, "name", name of boundDoc}
set end of nativeRows to {"doc", 0, "saved", saved of boundDoc}
set end of nativeRows to {"doc", 0, "path", posix full name of boundDoc as text}
set end of nativeRows to {"doc", 0, "counts.paragraphs", count paragraphs of boundDoc}
set end of nativeRows to {"doc", 0, "counts.tables", count tables of boundDoc}
set end of nativeRows to {"doc", 0, "counts.shapes", count shapes of boundDoc}
set end of nativeRows to {"doc", 0, "counts.sections", count sections of boundDoc}
set end of nativeRows to {"doc", 0, "counts.inline_shapes", count inline shapes of boundDoc}
set end of nativeRows to {"doc", 0, "counts.styles", count Word styles of boundDoc}
set objectCount to count paragraphs of boundDoc
repeat with i from 1 to objectCount
set targetObject to paragraph i of boundDoc
set targetRange to text object of targetObject
set end of nativeRows to {"paragraphs", i, "start", start of content of targetRange}
set end of nativeRows to {"paragraphs", i, "end", end of content of targetRange}
set end of nativeRows to {"paragraphs", i, "text", content of targetRange}
try
set end of nativeRows to {"paragraphs", i, "font.name", name of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "font.size", font size of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "font.bold", bold of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "font.italic", italic of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "font.underline", my enumIndex(underline of font object of targetRange, {underline none, underline single})}
end try
try
set end of nativeRows to {"paragraphs", i, "font.strikethrough", strike through of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "font.color", color of font object of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.alignment", my enumIndex(alignment of paragraph format of targetRange, {align paragraph left, align paragraph center, align paragraph right, align paragraph justify, align paragraph distribute})}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.left_indent", paragraph format left indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.right_indent", paragraph format right indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.first_line_indent", first line indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.space_before", space before of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.space_after", space after of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.line_spacing", line spacing of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.line_spacing_rule", my enumIndex(line spacing rule of paragraph format of targetRange, {line space single, line space1 pt5, line space double, line space at least, line space exactly, line space multiple})}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.keep_together", keep together of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.keep_with_next", keep with next of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.page_break_before", page break before of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "paragraph.widow_control", widow control of paragraph format of targetRange}
end try
try
set end of nativeRows to {"paragraphs", i, "style", name local of style of targetRange}
end try
end repeat
set objectCount to count tables of boundDoc
repeat with i from 1 to objectCount
set targetObject to table i of boundDoc
set end of nativeRows to {"tables", i, "rows", number of rows of targetObject}
set end of nativeRows to {"tables", i, "columns", number of columns of targetObject}
set end of nativeRows to {"tables", i, "allow_autofit", allow auto fit of targetObject}
repeat with targetCell in cells of text object of targetObject
set cellKey to (i as text) & "," & (row index of targetCell as text) & "," & (column index of targetCell as text)
set targetRange to text object of targetCell
set end of nativeRows to {"cells", cellKey, "start", start of content of targetRange}
set end of nativeRows to {"cells", cellKey, "end", end of content of targetRange}
set end of nativeRows to {"cells", cellKey, "text", content of targetRange}
try
set end of nativeRows to {"cells", cellKey, "font.name", name of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "font.size", font size of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "font.bold", bold of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "font.italic", italic of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "font.underline", my enumIndex(underline of font object of targetRange, {underline none, underline single})}
end try
try
set end of nativeRows to {"cells", cellKey, "font.strikethrough", strike through of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "font.color", color of font object of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.alignment", my enumIndex(alignment of paragraph format of targetRange, {align paragraph left, align paragraph center, align paragraph right, align paragraph justify, align paragraph distribute})}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.left_indent", paragraph format left indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.right_indent", paragraph format right indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.first_line_indent", first line indent of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.space_before", space before of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.space_after", space after of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.line_spacing", line spacing of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.line_spacing_rule", my enumIndex(line spacing rule of paragraph format of targetRange, {line space single, line space1 pt5, line space double, line space at least, line space exactly, line space multiple})}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.keep_together", keep together of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.keep_with_next", keep with next of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.page_break_before", page break before of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "paragraph.widow_control", widow control of paragraph format of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "style", name local of style of targetRange}
end try
try
set end of nativeRows to {"cells", cellKey, "width", width of targetCell}
end try
try
set end of nativeRows to {"cells", cellKey, "fill.color", background pattern color of shading of targetCell}
end try
try
set end of nativeRows to {"cells", cellKey, "vertical_alignment", my enumIndex(vertical alignment of targetCell, {cell align vertical top, cell align vertical center, cell align vertical bottom})}
end try
end repeat
end repeat
set objectCount to count shapes of boundDoc
repeat with i from 1 to objectCount
set targetObject to shape i of boundDoc
set end of nativeRows to {"shapes", i, "name", name of targetObject}
try
set end of nativeRows to {"shapes", i, "type", my enumIndex(shape type of targetObject, {shape type unset, shape type auto, shape type callout, shape type chart, shape type comment, shape type free form, shape type group, shape type embedded OLE control, shape type form control, shape type line, shape type linked OLE object, shape type linked picture, shape type OLE control, shape type picture, shape type place holder, shape type word art, shape type media, shape type text box, shape type script anchor, shape type table, shape type canvas, shape type diagram, shape type ink, shape type ink comment, shape type smartart graphic, shape type slicer, shape type web video, shape type content application, shape type graphic, shape type linked graphic})}
end try
try
set end of nativeRows to {"shapes", i, "geometry.left", left position of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "geometry.top", top of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "geometry.width", width of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "geometry.height", height of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "geometry.rotation", rotation of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "wrap", my enumIndex(wrap type of wrap format of targetObject, {wrap square, wrap tight, wrap through, wrap none, wrap top bottom, wrap behind, wrap front, wrap inline})}
end try
try
set end of nativeRows to {"shapes", i, "fill.color", fore color of fill format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "fill.back_color", back color of fill format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "fill.visible", visible of fill format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "fill.transparency", transparency of fill format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "line.color", fore color of line format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "line.weight", weight of line format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "line.visible", visible of line format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "line.transparency", transparency of line format of targetObject}
end try
try
set end of nativeRows to {"shapes", i, "text", content of text range of text frame of targetObject}
end try
end repeat
set objectCount to count sections of boundDoc
repeat with i from 1 to objectCount
set targetObject to section i of boundDoc
try
set end of nativeRows to {"sections", i, "page_setup.orientation", my enumIndex(orientation of page setup of targetObject, {orient portrait, orient landscape})}
end try
try
set end of nativeRows to {"sections", i, "page_setup.page_width", page width of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.page_height", page height of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.top_margin", top margin of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.bottom_margin", bottom margin of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.left_margin", left margin of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.right_margin", right margin of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.header_distance", header distance of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.footer_distance", footer distance of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "page_setup.gutter", gutter of page setup of targetObject}
end try
try
set end of nativeRows to {"sections", i, "columns", count text columns of page setup of targetObject}
end try
end repeat
end tell
end timeout
return my jsonRows(nativeRows)
