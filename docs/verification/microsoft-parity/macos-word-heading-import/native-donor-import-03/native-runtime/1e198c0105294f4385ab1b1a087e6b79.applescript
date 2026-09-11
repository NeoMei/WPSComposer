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
set boundDoc to document "document-df0a67d7641d45ffad4a41460737798b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ecpdpeqd/document-df0a67d7641d45ffad4a41460737798b.docx" then error "WPSC_STALE_DOCUMENT"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set carrierRange to text object of paragraph 4 of boundDoc
if start of content of carrierRange is not 67 or end of content of carrierRange is not 68 or content of carrierRange is not return then error "WPSC_V3_EMPTY_CARRIER"
set controlStyle to Word style "WPSC Heading Clone" of boundDoc
set style of carrierRange to controlStyle
set freshInsert to create range boundDoc start 67 end 67
set content of freshInsert to "FRESH STYLE 中文 العربية 😀"
set nativeRows to {{"stage",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
