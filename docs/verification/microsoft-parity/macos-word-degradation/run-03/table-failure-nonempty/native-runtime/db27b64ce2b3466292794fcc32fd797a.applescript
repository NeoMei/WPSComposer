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
set boundDoc to document "document-023aa3676a2f4caa803721d71fbe42ea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ftr0g2us/document-023aa3676a2f4caa803721d71fbe42ea.docx" then error "WPSC_STALE_DOCUMENT"
save as boundDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ftr0g2us/65d426ee2470471ba2b27c686f9b0fa2.pdf" file format format PDF add to recent files false
set boundDoc to document "document-023aa3676a2f4caa803721d71fbe42ea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ftr0g2us/document-023aa3676a2f4caa803721d71fbe42ea.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
