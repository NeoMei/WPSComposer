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
set boundDoc to document "document-e58ac26ad84345a3b49ac358b9301526.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ci3q2b/document-e58ac26ad84345a3b49ac358b9301526.docx" then error "WPSC_STALE_DOCUMENT"
set numberSelection to selection of boundWindow
if story type of numberSelection is not main text story then error "WPSC_NUMBERING_WRONG_STORY"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ci3q2b/document-e58ac26ad84345a3b49ac358b9301526.docx") then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ci3q2b/document-e58ac26ad84345a3b49ac358b9301526.docx") then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of numberSelection as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ci3q2b/document-e58ac26ad84345a3b49ac358b9301526.docx") then error "WPSC_STALE_DOCUMENT"
set numberSelectionRange to text object of numberSelection
if (start of content of numberSelectionRange) is not 73 or (end of content of numberSelectionRange) is not 73 then error "WPSC_NUMBERING_SELECTION_CHANGED"
set numberBefore to count fields of boundDoc
set numberRange to create range boundDoc start 73 end 73
create new field text range numberRange field type field style ref field text "\"标题 1\" \\s" preserve formatting true
set numberAfter to count fields of boundDoc
if numberAfter is not numberBefore + 1 then error "WPSC_NUMBERING_FIELD_DELTA"
set numberMatches to 0
set numberField to missing value
repeat with numberIndex from 1 to numberAfter
set numberCandidate to field numberIndex of boundDoc
if (start of content of field code of numberCandidate) is 74 then
set numberMatches to numberMatches + 1
set numberField to numberCandidate
end if
end repeat
if numberMatches is not 1 then error "WPSC_NUMBERING_FIELD_IDENTITY"
if field type of numberField is not field style ref then error "WPSC_NUMBERING_FIELD_TYPE"
set numberCode to content of field code of numberField as text
set numberCodeStart to start of content of field code of numberField
set numberCodeEnd to end of content of field code of numberField
set numberResultStart to start of content of result range of numberField
set numberResultEnd to end of content of result range of numberField
set numberNext to numberResultEnd + 1
if numberNext > (end of content of text object of boundDoc) - 1 then error "WPSC_NUMBERING_FIELD_RANGE"
make new bookmark at boundDoc with properties {name:"WPSC_N_4be5103a641f4ba987d6f19b9dbb85",text object:field code of numberField}
set numberIdentity to text object of bookmark "WPSC_N_4be5103a641f4ba987d6f19b9dbb85" of boundDoc
set selection start of numberSelection to numberNext
set selection end of numberSelection to numberNext
set nativeRows to {{"field",numberBefore,numberAfter,numberCodeStart,numberCodeEnd,numberResultStart,numberResultEnd,numberCode,"WPSC_N_4be5103a641f4ba987d6f19b9dbb85",start of content of numberIdentity,end of content of numberIdentity,content of numberIdentity as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
