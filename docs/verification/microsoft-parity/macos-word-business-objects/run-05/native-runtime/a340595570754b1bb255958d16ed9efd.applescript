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
set boundDoc to document "document-77a00d3c690d42fcb60cd1733d903bea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-e2k8a_fg/document-77a00d3c690d42fcb60cd1733d903bea.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set previousShapeCount to count shapes of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set insertionRange to create range boundDoc start insertionPoint end insertionPoint
set insertedShape to make new text box at boundDoc with properties {anchor:insertionRange, left position:72, top:360, width:220, height:54}
set shapeIndex to count shapes of boundDoc
set insertedShape to shape shapeIndex of boundDoc
set relative horizontal position of insertedShape to relative horizontal position page
set relative vertical position of insertedShape to relative vertical position page
set left position of insertedShape to 72
set top of insertedShape to 360
set width of insertedShape to 220
set height of insertedShape to 54
set content of text range of text frame of insertedShape to "Textbox 中文😀"
set font size of font object of text range of text frame of insertedShape to 13
set bold of font object of text range of text frame of insertedShape to true
set wrap type of wrap format of insertedShape to wrap square
set visible of fill format of insertedShape to true
set fore color of fill format of insertedShape to {65535, 61680, 46260}
if shapeIndex is not previousShapeCount + 1 then error "WPSC_TEXTBOX_COUNT_DELTA_FAILED"
set nativeRows to {{"created", "shape", shapeIndex, previousShapeCount, (shape type of insertedShape is shape type text box), content of text range of text frame of insertedShape as text, left position of insertedShape, top of insertedShape, width of insertedShape, height of insertedShape, my enumIndex(wrap type of wrap format of insertedShape, {wrap square, wrap tight, wrap through, wrap none, wrap top bottom, wrap behind, wrap front, wrap inline}), font size of font object of text range of text frame of insertedShape, bold of font object of text range of text frame of insertedShape, visible of fill format of insertedShape, fore color of fill format of insertedShape}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
