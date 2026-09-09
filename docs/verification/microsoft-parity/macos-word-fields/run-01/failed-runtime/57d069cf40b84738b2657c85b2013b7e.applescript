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
set boundDoc to document "document-b8a910f14ce24d9ea59a26a5fb149c73.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-vrlt1bll/document-b8a910f14ce24d9ea59a26a5fb149c73.docx" then error "WPSC_STALE_DOCUMENT"
repaginate boundDoc
repeat with storyIndex from 1 to count story ranges of boundDoc
set ownStory to story range storyIndex of boundDoc
set chainIndex to 0
repeat while ownStory is not missing value
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:" & (story type of ownStory as text) & "/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field style ref or field type of ownField is field sequence then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set ownStory to next story range of ownStory
end repeat
end repeat
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
