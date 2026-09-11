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
set boundDoc to document "document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_WINDOW_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档96"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档96") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 4a67af185bff49258c135ab74e46511c" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
activate object boundWindow
set selection start of selection of boundWindow to 12
set selection end of selection of boundWindow to 19
set diagActiveOwned to ((current application's NSString's stringWithString:(posix full name of document of active window as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") as boolean
if diagActiveOwned is not true then error "DIAG_ACTIVE_OWNER_CHANGED"
set diagSelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of document of diagSelection as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "DIAG_SELECTION_OWNER_CHANGED"
if story type of diagSelection is not main text story then error "DIAG_STORY_CHANGED"
set diagRange to create range boundDoc start 12 end 19
if start of content of diagRange is not 12 or end of content of diagRange is not 19 then error "DIAG_RANGE_CHANGED"
if (count tables of boundDoc) is not 1 then error "DIAG_SEED_CHANGED"
set diagPrefix to create range boundDoc start 0 end 12
set diagFullMarker to create range boundDoc start 12 end 19
if (content of diagFullMarker as text) is not "REPLACE" then error "DIAG_MARKER_CHANGED"
set nativeRows to {{"before",posix full name of boundDoc as text,12,end of content of diagRange,count tables of boundDoc,content of text object of boundDoc as text,content of text object of bookmark "semantic_old_table" of boundDoc as text,content of diagPrefix as text,content of text object of bookmark "semantic_suffix" of boundDoc as text,saved of boundDoc,name of document of active window as text,diagActiveOwned}}
set diagResult to "created"
set diagCode to 0
set diagMessage to ""
set content of diagRange to "WPSC_R1C1" & tab & "WPSC_R1C2" & tab & "WPSC_R1C3" & return & "WPSC_R2C1" & tab & "WPSC_R2C2" & tab & "WPSC_R2C3" & return & "WPSC_R3C1" & tab & "WPSC_R3C2" & tab & "WPSC_R3C3" & return & "WPSC_R4C1" & tab & "WPSC_R4C2" & tab & "WPSC_R4C3"
set diagRange to create range boundDoc start 12 end 131
if start of content of diagRange is not 12 or end of content of diagRange is not 131 then error "GRID_RANGE_CHANGED"
set gridBeforeText to content of diagRange as text
if not ((current application's NSString's stringWithString:(gridBeforeText))'s isEqualToString:"WPSC_R1C1" & tab & "WPSC_R1C2" & tab & "WPSC_R1C3" & return & "WPSC_R2C1" & tab & "WPSC_R2C2" & tab & "WPSC_R2C3" & return & "WPSC_R3C1" & tab & "WPSC_R3C2" & tab & "WPSC_R3C3" & return & "WPSC_R4C1" & tab & "WPSC_R4C2" & tab & "WPSC_R4C3") then error "GRID_PREPOPULATION_MISMATCH"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档96"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档96") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 4a67af185bff49258c135ab74e46511c" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
set end of nativeRows to {"grid",12,131,gridBeforeText}
set diagTable to convert to table diagRange separator separate by tabs number of rows 4 number of columns 3
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-glg0bgrl/document-14d4fd91b1ce4cf3aa46a244cf2d6584.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档96"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档96") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 4a67af185bff49258c135ab74e46511c" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
set end of nativeRows to {"attempt",diagResult,diagCode,diagMessage}
set end of nativeRows to {"after",count tables of boundDoc,content of text object of boundDoc as text,saved of boundDoc}
repeat with diagIndex from 1 to (count tables of boundDoc)
set diagObserved to table diagIndex of boundDoc
set end of nativeRows to {"table",diagIndex,start of content of text object of diagObserved,end of content of text object of diagObserved,number of rows of diagObserved,number of columns of diagObserved,content of text object of diagObserved as text}
end repeat
if diagResult is "created" then
set conversionStart to start of content of text object of diagTable
set conversionEnd to end of content of text object of diagTable
set conversionBefore to create range boundDoc start 0 end conversionStart
set conversionAfter to create range boundDoc start conversionEnd end (end of content of text object of boundDoc)
set end of nativeRows to {"conversion",true,conversionStart,conversionEnd,number of rows of diagTable,number of columns of diagTable,content of text object of diagTable as text,content of conversionBefore as text,content of conversionAfter as text}
else
set end of nativeRows to {"conversion",false,-1,-1,0,0,"","",""}
end if
set conversionPrefix to create range boundDoc start 0 end 12
set end of nativeRows to {"surroundings",content of conversionPrefix as text,content of text object of bookmark "semantic_suffix" of boundDoc as text,content of text object of bookmark "semantic_old_table" of boundDoc as text}
set gridCell to get cell from table diagTable row 1 column 1
set end of nativeRows to {"grid-cell",1,1,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 1 column 2
set end of nativeRows to {"grid-cell",1,2,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 1 column 3
set end of nativeRows to {"grid-cell",1,3,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 2 column 1
set end of nativeRows to {"grid-cell",2,1,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 2 column 2
set end of nativeRows to {"grid-cell",2,2,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 2 column 3
set end of nativeRows to {"grid-cell",2,3,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 3 column 1
set end of nativeRows to {"grid-cell",3,1,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 3 column 2
set end of nativeRows to {"grid-cell",3,2,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 3 column 3
set end of nativeRows to {"grid-cell",3,3,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 4 column 1
set end of nativeRows to {"grid-cell",4,1,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 4 column 2
set end of nativeRows to {"grid-cell",4,2,content of text object of gridCell as text}
set gridCell to get cell from table diagTable row 4 column 3
set end of nativeRows to {"grid-cell",4,3,content of text object of gridCell as text}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
