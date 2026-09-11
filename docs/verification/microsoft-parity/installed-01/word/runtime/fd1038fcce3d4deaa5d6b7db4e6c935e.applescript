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
set boundDoc to document "document-0a6eaa69515a400b95aa275313ae2114.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-4xechqj9/document-0a6eaa69515a400b95aa275313ae2114.docx" then error "WPSC_STALE_DOCUMENT"
set targetRange to text object of paragraph 1 of boundDoc
set font size of font object of targetRange to 16
set bold of font object of targetRange to true
set replacementStart to start of content of targetRange
set replacementEnd to end of content of targetRange
set replacementRange to create range boundDoc start replacementStart end (replacementEnd - 1)
set content of replacementRange to "Installed native Word"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
