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
set boundDoc to document "document-56462579dcb84c729ce1a2b6c7b5bb50.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fyj72s8k/document-56462579dcb84c729ce1a2b6c7b5bb50.docx" then error "WPSC_STALE_DOCUMENT"
set paginationText to content of text object of boundDoc as text
set paginationHash to do shell script ("/usr/bin/printf %s " & quoted form of paginationText & " | /usr/bin/shasum -a 256")
set nativeRows to {{"state",paginationHash,end of content of text object of boundDoc,count paragraphs of boundDoc,count tables of boundDoc,count bookmarks of boundDoc,count fields of boundDoc,saved of boundDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
