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
set boundDoc to document "document-210b8a8e1720419d900bae5a87b2fcd3.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-bq2enbla/document-210b8a8e1720419d900bae5a87b2fcd3.docx" then error "WPSC_STALE_DOCUMENT"
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
set content of semanticRange to "Figure 😀" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 10)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field sequence field text "WPSC_FIG" preserve formatting true
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
set content of r to " Caption 😀" & return
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
set content of semanticRange to "Table 😀" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 9)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field sequence field text "WPSC_TAB" preserve formatting true
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
set content of r to " Caption 😀" & return
set r to create range boundDoc start 0 end 8
make new bookmark at boundDoc with properties {name:"WPSC_RefProbe",text object:r}
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field ref field text "WPSC_RefProbe" preserve formatting true
set insertionPoint to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start insertionPoint end insertionPoint
set content of r to return
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
