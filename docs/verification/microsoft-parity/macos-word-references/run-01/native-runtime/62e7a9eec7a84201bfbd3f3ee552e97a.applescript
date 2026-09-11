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
set boundDoc to document "document-e099edb5171c4eca8b004158045dc5e3.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-pdx7rs5d/document-e099edb5171c4eca8b004158045dc5e3.docx" then error "WPSC_STALE_DOCUMENT"
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
set content of paragraphRange to "[3] 中文😀 已引" & return & ""
set insertionPoint to insertionPoint + 12
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set alignment of paragraph format of paragraphRange to align paragraph left
set paragraph format left indent of paragraph format of paragraphRange to 21.5
set first line indent of paragraph format of paragraphRange to -12.25
set space before of paragraph format of paragraphRange to 0
set space after of paragraph format of paragraphRange to 4.5
set keep together of paragraph format of paragraphRange to true
set end of nativeRows to {"paragraph",0,paragraphStart,insertionPoint,content of paragraphRange as text,my enumIndex(alignment of paragraph format of paragraphRange,{align paragraph left}),paragraph format left indent of paragraph format of paragraphRange,first line indent of paragraph format of paragraphRange,space before of paragraph format of paragraphRange,space after of paragraph format of paragraphRange,keep together of paragraph format of paragraphRange}
set paragraphStart to insertionPoint
set paragraphRange to create range boundDoc start insertionPoint end insertionPoint
set content of paragraphRange to "[9] 未引" & return & ""
set insertionPoint to insertionPoint + 7
set paragraphRange to create range boundDoc start paragraphStart end insertionPoint
set alignment of paragraph format of paragraphRange to align paragraph left
set paragraph format left indent of paragraph format of paragraphRange to 21.5
set first line indent of paragraph format of paragraphRange to -12.25
set space before of paragraph format of paragraphRange to 0
set space after of paragraph format of paragraphRange to 4.5
set keep together of paragraph format of paragraphRange to true
set end of nativeRows to {"paragraph",1,paragraphStart,insertionPoint,content of paragraphRange as text,my enumIndex(alignment of paragraph format of paragraphRange,{align paragraph left}),paragraph format left indent of paragraph format of paragraphRange,first line indent of paragraph format of paragraphRange,space before of paragraph format of paragraphRange,space after of paragraph format of paragraphRange,keep together of paragraph format of paragraphRange}
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
