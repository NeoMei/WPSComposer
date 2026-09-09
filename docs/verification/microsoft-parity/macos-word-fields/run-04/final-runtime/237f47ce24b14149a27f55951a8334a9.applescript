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
set boundDoc to document "document-ee73e5a3ccac4affb323fb8c1a7bb9da.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-slctmwwj/document-ee73e5a3ccac4affb323fb8c1a7bb9da.docx" then error "WPSC_STALE_DOCUMENT"
repeat with tocIndex from 1 to (count of tables of contents of boundDoc)
update (table of contents tocIndex of boundDoc)
end repeat
repeat with figureIndex from 1 to (count of tables of figures of boundDoc)
update (table of figures figureIndex of boundDoc)
end repeat
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
