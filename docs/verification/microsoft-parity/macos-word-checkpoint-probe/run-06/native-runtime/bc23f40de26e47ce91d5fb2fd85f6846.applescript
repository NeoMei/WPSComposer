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
set boundDoc to document "document-96bfb25a4cb14857b20b9cfae7d1f5ad.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9eizwnep/document-96bfb25a4cb14857b20b9cfae7d1f5ad.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set insertionPoint to (end of content of text object of boundDoc) - 1
set insertionRange to create range boundDoc start insertionPoint end insertionPoint
set insertedTable to make new table at boundDoc with properties {text object:insertionRange,number of rows:2,number of columns:2}
set content of text object of cell 1 of row 1 of insertedTable to "T1"
set content of text object of cell 2 of row 1 of insertedTable to "表😀"
set nativeRows to {{"table-insert-ack",start of content of text object of insertedTable,end of content of text object of insertedTable}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
