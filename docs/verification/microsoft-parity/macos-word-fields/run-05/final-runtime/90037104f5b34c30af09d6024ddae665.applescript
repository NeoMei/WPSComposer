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
set boundDoc to document "document-de797c88dbce4c9c90073173f3a023ec.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-dvgjer62/document-de797c88dbce4c9c90073173f3a023ec.docx" then error "WPSC_STALE_DOCUMENT"
repaginate boundDoc
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type main text story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:main text/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnotes story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnotes story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type comments story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:comments/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type text frame story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:text frame/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote separator story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 13
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote continuation notice story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote separator story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 16
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote continuation notice story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
if field type of ownField is field page or field type of ownField is field num pages then
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
