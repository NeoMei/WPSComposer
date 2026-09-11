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
set boundDoc to document "document-81ce5fa3221f4b7aba765048252f5f1b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-0tgbyg94/document-81ce5fa3221f4b7aba765048252f5f1b.docx" then error "WPSC_STALE_DOCUMENT"
set p to (end of content of text object of boundDoc) - 1
set insertionRange to create range boundDoc start p end p
set beforeCount to count inline shapes of boundDoc
make new standard inline horizontal line at insertionRange
if (count inline shapes of boundDoc) is not beforeCount + 1 then error "WPSC_RULE_COUNT_DELTA_FAILED"
set ownLine to inline shape 1 of boundDoc
set nativeRows to {{"inline",count inline shapes of boundDoc,(inline shape type of ownLine is inline shape horizontal line),width of ownLine,height of ownLine}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
