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
set boundDoc to document "document-7c6ba1877ddd4683913fa03a22f2b578.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ly1lxrfe/document-7c6ba1877ddd4683913fa03a22f2b578.docx" then error "WPSC_STALE_DOCUMENT"
set roleName to "WpsComposerSectionRole_" & ((count sections of boundDoc) as text)
set roleMatches to {}
repeat with roleIndex from 1 to (count variables of boundDoc)
set candidateRole to variable roleIndex of boundDoc
if (name of candidateRole as text) is roleName then set end of roleMatches to candidateRole
end repeat
if (count roleMatches) > 1 then error "WPSC_ROLE_AMBIGUOUS"
if (count roleMatches) is 0 then
make new variable at boundDoc with properties {name:roleName, variable value:"media😀"}
end if
set ownRole to variable roleName of boundDoc
set variable value of ownRole to "media😀"
if (variable value of ownRole as text) is not "media😀" then error "WPSC_ROLE_READBACK_FAILED"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
