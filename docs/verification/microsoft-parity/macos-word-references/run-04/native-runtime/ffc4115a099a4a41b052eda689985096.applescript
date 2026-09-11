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
set boundDoc to document "document-62b78f798da34e6285d912d125996c1b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-owklc5il/document-62b78f798da34e6285d912d125996c1b.docx" then error "WPSC_STALE_DOCUMENT"
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
set paragraphStart to insertionPoint
set paragraphRange to create range boundDoc start insertionPoint end insertionPoint
set content of paragraphRange to "[8] 原样条目" & return & ""
set insertionPoint to insertionPoint + 9
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set end of nativeRows to {"paragraph",0,paragraphStart,insertionPoint,content of paragraphRange as text}
set paragraphStart to insertionPoint
set paragraphRange to create range boundDoc start insertionPoint end insertionPoint
set content of paragraphRange to "7" & return & ""
set insertionPoint to insertionPoint + 2
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set end of nativeRows to {"paragraph",1,paragraphStart,insertionPoint,content of paragraphRange as text}
set paragraphStart to insertionPoint
set paragraphRange to create range boundDoc start insertionPoint end insertionPoint
set content of paragraphRange to "False" & return & ""
set insertionPoint to insertionPoint + 6
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set end of nativeRows to {"paragraph",2,paragraphStart,insertionPoint,content of paragraphRange as text}
set paragraphStart to insertionPoint
set paragraphRange to create range boundDoc start insertionPoint end insertionPoint
set content of paragraphRange to "None" & return & ""
set insertionPoint to insertionPoint + 5
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set end of nativeRows to {"paragraph",3,paragraphStart,insertionPoint,content of paragraphRange as text}
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
