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
set boundDoc to document "document-36f80d459f444dceb18d337207eb9e19.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-gacvl1bk/document-36f80d459f444dceb18d337207eb9e19.docx" then error "WPSC_STALE_DOCUMENT"
set semanticStyle to Word style (style heading2) of boundDoc
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 14
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 15
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
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
set content of semanticRange to "Direct Heading Two" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 19)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set outline level of paragraph format of semanticRange to outline level2
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
