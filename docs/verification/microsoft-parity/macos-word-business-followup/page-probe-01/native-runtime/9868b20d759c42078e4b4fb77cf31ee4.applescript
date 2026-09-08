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
set boundDoc to document "document-2b4d15e314c148b89dbe7d85f769e9c1.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jyaxeh1u/document-2b4d15e314c148b89dbe7d85f769e9c1.docx" then error "WPSC_STALE_DOCUMENT"
set pagePart to get footer (section 1 of boundDoc) index header footer primary
set content of text object of pagePart to "Page "
make new page number at pagePart with properties {alignment:align page number left}
set nativeRows to {{content of text object of boundDoc,content of text object of pagePart}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
