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
set boundDoc to document "document-2513ae7dbb9a4ed89ec7ac59e9ba45a0.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-3mxu8geh/document-2513ae7dbb9a4ed89ec7ac59e9ba45a0.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-3mxu8geh/document-2513ae7dbb9a4ed89ec7ac59e9ba45a0.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-3mxu8geh/document-2513ae7dbb9a4ed89ec7ac59e9ba45a0.docx") then error "QUALITY_WINDOW_CHANGED"
set p to (end of content of text object of boundDoc) - 1
set insertionRange to create range boundDoc start p end p
set beforeCount to count inline shapes of boundDoc
set creationResult to make new standard inline horizontal line at boundDoc with properties {text object:insertionRange}
set totalInlineAfter to count inline shapes of boundDoc
set ownedCharacterCount to count characters of text object of boundDoc
set standardCountAfter to -1
set standardCountError to ""
try
set standardCountAfter to count standard inline horizontal lines of boundDoc
on error diagnosticMessage number diagnosticNumber
set standardCountError to (diagnosticNumber as text) & ":" & diagnosticMessage
end try
log my jsonRows({{"inline-creation-counts",beforeCount,totalInlineAfter,standardCountAfter,ownedCharacterCount,standardCountError}})
try
log {"inline-creation-result-type",class of creationResult}
on error diagnosticMessage number diagnosticNumber
log {"inline-creation-result-type-error",diagnosticNumber,diagnosticMessage}
end try
if (count inline shapes of boundDoc) is not beforeCount + 1 then error "WPSC_RULE_COUNT_DELTA_FAILED"
set ownLine to inline shape 1 of boundDoc
set nativeRows to {{"inline",count inline shapes of boundDoc,(inline shape type of ownLine is inline shape horizontal line),width of ownLine,height of ownLine}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
