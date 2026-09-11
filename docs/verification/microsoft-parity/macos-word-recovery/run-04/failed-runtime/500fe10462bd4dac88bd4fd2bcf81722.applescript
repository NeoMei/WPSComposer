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
set boundDoc to document "document-210ca8985c0942949d66360c6de56031.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-zqca0bbu/document-210ca8985c0942949d66360c6de56031.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set operationStart to insertionPoint
set beforeParagraphs to count paragraphs of boundDoc
set nativeRows to {}
set referenceFailed to false
set styleRanges to {}
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "FALLBACK-ONCE"
set insertionPoint to insertionPoint + 13
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "FALLBACK-ONCE" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "" & return & ""
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "" & return & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
repeat with styleIndex from 1 to count styleRanges
set styleSpec to item styleIndex of styleRanges
set styledRange to create range boundDoc start (item 2 of styleSpec) end (item 3 of styleSpec)
set italic of font object of styledRange to true
set color of font object of styledRange to {40092, 0, 1542}
set background pattern color of shading of styledRange to {64764, 59624, 59110}
set end of nativeRows to {"styled",item 1 of styleSpec,start of content of styledRange,end of content of styledRange,italic of font object of styledRange,(color of font object of styledRange is {40092, 0, 1542}),(background pattern color of shading of styledRange is {64764, 59624, 59110})}
end repeat
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end if
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
