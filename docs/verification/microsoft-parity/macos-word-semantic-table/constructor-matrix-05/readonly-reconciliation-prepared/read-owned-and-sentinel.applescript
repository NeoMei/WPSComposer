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
if ((current application's NSString's stringWithString:(posix full name of reconciliationCandidate as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then
set reconciliationMatches to reconciliationMatches + 1
set reconciliationDoc to reconciliationCandidate
end if
end repeat
if reconciliationMatches is not 1 then error "DIAG_OWNED_IDENTITY_UNVERIFIED"
set qualitySentinel to document "文档96"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档96") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 4a67af185bff49258c135ab74e46511c" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(posix full name of reconciliationDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "DIAG_OWNED_PATH_CHANGED"
set nativeRows to {{"binding",posix full name of reconciliationDoc as text,name of qualitySentinel as text},{"body",end of content of text object of reconciliationDoc,content of text object of reconciliationDoc as text,count tables of reconciliationDoc}}
repeat with reconciliationIndex from 1 to (count tables of reconciliationDoc)
set reconciliationTable to table reconciliationIndex of reconciliationDoc
set reconciliationRange to text object of reconciliationTable
set end of nativeRows to {"table",reconciliationIndex,start of content of reconciliationRange,end of content of reconciliationRange,number of rows of reconciliationTable,number of columns of reconciliationTable,content of reconciliationRange as text}
end repeat
end tell
return my jsonRows(nativeRows)
