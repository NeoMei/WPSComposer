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
set boundDoc to document "document-c86c99aae54242b3b0716ac26bfcaaf5.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-iew3av7r/document-c86c99aae54242b3b0716ac26bfcaaf5.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
set expectedBodyStyleName to name local of (Word style (style body text) of boundDoc) as text
set expectedListStyleName to name local of (Word style (style list paragraph) of boundDoc) as text
set end of nativeRows to {"builtin-styles",expectedBodyStyleName,expectedListStyleName}
repeat with paragraphIndex from 1 to count paragraphs of boundDoc
set nativeParagraph to text object of paragraph paragraphIndex of boundDoc
set nativeFormat to paragraph format of nativeParagraph
set nativeStyleName to name local of style of nativeParagraph as text
set end of nativeRows to {paragraphIndex as integer,content of nativeParagraph as text,nativeStyleName,paragraph format left indent of nativeFormat,first line indent of nativeFormat}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
