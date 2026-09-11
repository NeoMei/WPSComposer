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
set boundDoc to document "document-71675695ba8e4e8b9b04b0c35fa4e6fd.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-unydvra4/document-71675695ba8e4e8b9b04b0c35fa4e6fd.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
set ownRange to create range boundDoc start insertionPoint end insertionPoint
create new field text range ownRange field type field ref field text "wpsc_checkpoint_preexisting \\h" preserve formatting true
set ownField to field (count fields of boundDoc) of boundDoc
if (update field ownField) is false then error "FIELD_UPDATE_FAILED"
set nativeRows to {{"append-ref-ack",count fields of boundDoc,content of field code of ownField as text,start of content of result range of ownField,end of content of result range of ownField}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
