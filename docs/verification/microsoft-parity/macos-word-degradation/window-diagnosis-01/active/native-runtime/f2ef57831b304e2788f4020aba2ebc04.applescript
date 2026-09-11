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
set boundDoc to document "document-377a3cce318a43b6ae6ecd65915c508a.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-azd0gerp/document-377a3cce318a43b6ae6ecd65915c508a.docx" then error "WPSC_STALE_DOCUMENT"
set beforeText to content of text object of boundDoc as text
set beforeEnd to end of content of text object of boundDoc
set beforeCount to count tables of boundDoc
set probeResult to {"not-run"}
try
activate object boundWindow
set p to (end of content of text object of boundDoc) - 1
set r to create range boundDoc start p end p
set t to make new table at boundDoc with properties {text object:r,number of rows:1,number of columns:1}
set probeResult to {"created",start of content of text object of t,end of content of text object of t,count rows of t,count columns of t}
on error probeError number probeNumber
set probeResult to {"error",probeNumber,probeError}
end try
set nativeRows to {{"before",beforeText,beforeEnd,beforeCount},probeResult,{"after",content of text object of boundDoc as text,end of content of text object of boundDoc,count tables of boundDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
