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
set boundDoc to document "document-de797c88dbce4c9c90073173f3a023ec.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-dvgjer62/document-de797c88dbce4c9c90073173f3a023ec.docx" then error "WPSC_STALE_DOCUMENT"
set semanticStyle to Word style (style normal) of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set semanticRange to create range boundDoc start insertionPoint end insertionPoint
set content of semanticRange to "Literal figures" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 16)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set font size of font object of semanticRange to 14
set bold of font object of semanticRange to true
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set semanticStyle to Word style (style body text) of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set semanticRange to create range boundDoc start insertionPoint end insertionPoint
set content of semanticRange to "[Figure index placeholder]" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 27)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
