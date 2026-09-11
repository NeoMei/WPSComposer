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
repeat with i from (count of documents) to 1 by -1
set d to document i
if posix full name of d is "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-m18pihz8/document.docx" then
close d saving no
else if (name of d is "文档54") and (content of text object of d is "WPSC-SENTINEL-87048479e59d401398167c04a914a912" & return) then
if saved of d then error "WPSC_SENTINEL_SAVED"
close d saving no
end if
end repeat
set nativeRows to {{"clean"}}
end tell
end timeout
return my jsonRows(nativeRows)
