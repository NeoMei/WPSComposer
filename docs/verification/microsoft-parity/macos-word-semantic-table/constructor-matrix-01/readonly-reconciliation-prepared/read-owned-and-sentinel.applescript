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
if not application "Microsoft Word" is running then error "DIAG_WORD_NOT_RUNNING"
tell application "/Applications/Microsoft Word.app"
set reconciliationMatches to 0
repeat with reconciliationIndex from 1 to (count documents)
set reconciliationCandidate to document reconciliationIndex
if ((current application's NSString's stringWithString:(posix full name of reconciliationCandidate as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-g6k8gzpf/document-798995b924bf4696abfc488e45e139b6.docx") then
set reconciliationMatches to reconciliationMatches + 1
set reconciliationDoc to reconciliationCandidate
end if
end repeat
if reconciliationMatches is not 1 then error "DIAG_OWNED_IDENTITY_UNVERIFIED"
set qualitySentinel to document "文档92"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档92") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 da007ee4a6e74be6b7a5e39b76e7c136" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(posix full name of reconciliationDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-g6k8gzpf/document-798995b924bf4696abfc488e45e139b6.docx") then error "DIAG_OWNED_PATH_CHANGED"
set nativeRows to {{"binding",posix full name of reconciliationDoc as text,name of qualitySentinel as text},{"body",end of content of text object of reconciliationDoc,content of text object of reconciliationDoc as text,count tables of reconciliationDoc}}
repeat with reconciliationIndex from 1 to (count tables of reconciliationDoc)
set reconciliationTable to table reconciliationIndex of reconciliationDoc
set reconciliationRange to text object of reconciliationTable
set end of nativeRows to {"table",reconciliationIndex,start of content of reconciliationRange,end of content of reconciliationRange,number of rows of reconciliationTable,number of columns of reconciliationTable,content of reconciliationRange as text}
end repeat
end tell
return my jsonRows(nativeRows)
