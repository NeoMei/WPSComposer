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
set boundDoc to document "document-bbcdd06a9c974e52bb2b0d3f0da20bbb.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-txw2p1e4/document-bbcdd06a9c974e52bb2b0d3f0da20bbb.docx" then error "WPSC_STALE_DOCUMENT"
set sourceStyle to Word style (style heading1) of boundDoc
set linkedStyle to Word style "标题 1 字符" of boundDoc
set cloneStyle to Word style "WPSC Heading Clone" of boundDoc
set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc,italic of font object of sourceStyle,italic of font object of linkedStyle,italic of font object of cloneStyle}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
