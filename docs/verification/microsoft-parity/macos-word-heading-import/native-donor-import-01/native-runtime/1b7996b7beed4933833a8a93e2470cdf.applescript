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
set boundDoc to document "document-bc0a65e8ead24997bf64e4c3ee7cef55.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-4tsnpylb/document-bc0a65e8ead24997bf64e4c3ee7cef55.docx" then error "WPSC_STALE_DOCUMENT"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "REPLACE" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set clearRange to create range boundDoc start 66 end 74
if content of clearRange is not "REPLACE" & return then error "WPSC_COLLAPSED_TARGET_CHANGED"
set clearStart to start of content of clearRange
set clearEnd to end of content of clearRange
if clearStart is not 66 or clearEnd is not 74 then error "WPSC_COLLAPSED_TARGET_BOUNDS"
set content of clearRange to ""
set nativeRows to {{"clear",clearStart,clearEnd,content of text object of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
