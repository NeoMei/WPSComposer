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
set boundDoc to document "document-3e8ba73679464786b4aa35ba0a39180e.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-f0jk6515/document-3e8ba73679464786b4aa35ba0a39180e.docx" then error "WPSC_STALE_DOCUMENT"
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
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "Static "
set insertionPoint to insertionPoint + 7
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "Static " then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "("
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "(" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "FALLBACK-ONLY"
set insertionPoint to insertionPoint + 13
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "FALLBACK-ONLY" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
set italic of font object of literalRange to true
set color of font object of literalRange to {40092, 0, 1542}
set background pattern color of shading of literalRange to {64764, 59624, 59110}
set end of nativeRows to {"styled",1,literalStart,insertionPoint,italic of font object of literalRange,(color of font object of literalRange is {40092, 0, 1542}),(background pattern color of shading of literalRange is {64764, 59624, 59110})}
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to ")"
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not ")" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "" & return & ""
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "" & return & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",2,literalStart,insertionPoint,content of literalRange as text}
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end if
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
