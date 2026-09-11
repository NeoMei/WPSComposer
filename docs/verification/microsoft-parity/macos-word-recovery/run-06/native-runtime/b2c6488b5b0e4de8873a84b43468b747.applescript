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
set boundDoc to document "document-4752f6ea6abc4ef082cd860b1fc94b54.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9ze4opyl/document-4752f6ea6abc4ef082cd860b1fc94b54.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set ownTable to make new table at boundDoc with properties {text object:r,number of rows:1,number of columns:1}
set p to start of content of text object of cell 1 of row 1 of ownTable
set r to create range boundDoc start p end p
create new field text range r field type field ref field text "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa" preserve formatting true
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set content of r to "PARTIAL-OUTSIDE "
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
create new field text range r field type field ref field text "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa" preserve formatting true
set nativeRows to {{"partial-native-ack",count fields of boundDoc,count tables of boundDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
