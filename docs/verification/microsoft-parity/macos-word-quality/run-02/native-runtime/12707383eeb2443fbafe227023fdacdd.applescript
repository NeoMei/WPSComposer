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
-- WPSC_BIND_NEW
set previousNames to name of every document
set boundDoc to make new document
set newName to name of boundDoc
if previousNames contains newName then error "WPSC_NEW_IDENTITY_COLLISION"
set boundDoc to document newName
save as boundDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx" file format format document default add to recent files false
set boundDoc to document "document-db70cbc45b1443f09834b808a00b9462.docx"
if posix full name of boundDoc is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx" then error "WPSC_BINDING_FAILED"
set boundWindow to active window of boundDoc
set nativeRows to {{"binding", id of boundWindow, posix full name of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
