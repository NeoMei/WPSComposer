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
set boundDoc to document "document-23ed51f595644bca902e8076a0007ec9.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-b03g3kic/document-23ed51f595644bca902e8076a0007ec9.docx" then error "WPSC_STALE_DOCUMENT"
set importRange to text object of paragraph 2 of boundDoc
set importStart to start of content of importRange
set importEnd to end of content of importRange
if content of importRange is not "REPLACE" & return then error "WPSC_IMPORT_PREIMAGE"
insert file at importRange file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-b03g3kic/heading-input.xml" confirm conversions false link false
set nativeRows to {{"import",importStart,importEnd,content of text object of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
