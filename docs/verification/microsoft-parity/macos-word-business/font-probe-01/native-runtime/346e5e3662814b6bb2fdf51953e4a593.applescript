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
set boundDoc to document "document-1492bc22a90c41a3bdedef7976d51d9f.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5fir20uz/document-1492bc22a90c41a3bdedef7976d51d9f.docx" then error "WPSC_STALE_DOCUMENT"
set r to text object of paragraph 1 of boundDoc
set name of font object of r to "Consolas"
set ascii name of font object of r to "Consolas"
set other name of font object of r to "Consolas"
set complex script name of font object of r to "Consolas"
set nativeRows to {{name of font object of r, east asian name of font object of r,ascii name of font object of r,other name of font object of r,complex script name of font object of r}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
