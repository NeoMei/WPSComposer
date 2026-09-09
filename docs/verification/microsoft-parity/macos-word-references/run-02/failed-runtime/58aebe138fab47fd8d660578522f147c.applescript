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
set boundDoc to document "document-6fe260f50f6e4df181b7fcbc16f27abc.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-eddiheun/document-6fe260f50f6e4df181b7fcbc16f27abc.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field sequence field text "WPSC_FIG \\r 7" preserve formatting true
set ownField to field (count fields of boundDoc) of boundDoc
if (update field ownField) is false then error "TARGET_UPDATE_FAILED"
make new bookmark at boundDoc with properties {name:"wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa",text object:result range of ownField}
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set content of r to return
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field sequence field text "WPSC_TAB \\r 11" preserve formatting true
set ownField to field (count fields of boundDoc) of boundDoc
if (update field ownField) is false then error "TARGET_UPDATE_FAILED"
make new bookmark at boundDoc with properties {name:"wpsc_tab_bbbbbbbbbbbbbbbbbbbbbbbb",text object:result range of ownField}
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set content of r to return
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set r to create range boundDoc start insertionPoint end insertionPoint
create new field text range r field type field sequence field text "WPSC_EQ \\r 13" preserve formatting true
set ownField to field (count fields of boundDoc) of boundDoc
if (update field ownField) is false then error "TARGET_UPDATE_FAILED"
make new bookmark at boundDoc with properties {name:"wpsc_eq_cccccccccccccccccccccccc",text object:result range of ownField}
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set content of r to return
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
