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
set boundDoc to document "document-db70cbc45b1443f09834b808a00b9462.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx") then error "QUALITY_WINDOW_CHANGED"
activate object boundWindow
set content of text object of boundDoc to "PREFIX 中文😀" & return & "PREFIX FIELD " & return & "PREFIX TABLE SLOT" & return & "BEFORE BOUNDARY" & return & "AFTER BOUNDARY" & return & "SUFFIX FIELD " & return & "SUFFIX TABLE SLOT" & return & "TAIL 中文😀" & return & ""
set seedRange to text object of paragraph 1 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_1",text object:seedRange}
set seedRange to text object of paragraph 2 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_2",text object:seedRange}
set seedRange to text object of paragraph 3 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_3",text object:seedRange}
set seedRange to text object of paragraph 4 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_4",text object:seedRange}
set seedRange to text object of paragraph 5 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_5",text object:seedRange}
set seedRange to text object of paragraph 6 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_6",text object:seedRange}
set seedRange to text object of paragraph 7 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_7",text object:seedRange}
set seedRange to text object of paragraph 8 of boundDoc
make new bookmark at boundDoc with properties {name:"quality_seed_8",text object:seedRange}
set seedPoint to (end of content of text object of bookmark "quality_seed_6" of boundDoc) - 1
set seedRange to create range boundDoc start seedPoint end seedPoint
create new field text range seedRange field type field sequence field text "Quality6" preserve formatting true
set seedPoint to (end of content of text object of bookmark "quality_seed_2" of boundDoc) - 1
set seedRange to create range boundDoc start seedPoint end seedPoint
create new field text range seedRange field type field sequence field text "Quality2" preserve formatting true
set seedPoint to start of content of text object of bookmark "quality_seed_7" of boundDoc
set seedRange to create range boundDoc start seedPoint end seedPoint
set seedTable to make new table at boundDoc with properties {text object:seedRange,number of rows:1,number of columns:1}
set content of text object of (get cell from table seedTable row 1 column 1) to "EXISTING TABLE 7"
set seedPoint to start of content of text object of bookmark "quality_seed_3" of boundDoc
set seedRange to create range boundDoc start seedPoint end seedPoint
set seedTable to make new table at boundDoc with properties {text object:seedRange,number of rows:1,number of columns:1}
set content of text object of (get cell from table seedTable row 1 column 1) to "EXISTING TABLE 3"
set seedRange to text object of bookmark "quality_seed_4" of boundDoc
set seedFormat to paragraph format of seedRange
set first line indent of seedFormat to 17
set space before of seedFormat to 11
set space after of seedFormat to 7
set line spacing rule of seedFormat to line space exactly
set line spacing of seedFormat to 21
set keep with next of seedFormat to true
set widow control of seedFormat to false
set italic of font object of seedRange to true
set font size of font object of seedRange to 14
set seedRange to text object of bookmark "quality_seed_5" of boundDoc
set seedFormat to paragraph format of seedRange
set first line indent of seedFormat to 23
set space before of seedFormat to 13
set space after of seedFormat to 9
set line spacing rule of seedFormat to line space exactly
set line spacing of seedFormat to 21
set keep with next of seedFormat to true
set widow control of seedFormat to false
set italic of font object of seedRange to true
set font size of font object of seedRange to 14
set seedPoint to start of content of text object of bookmark "quality_seed_5" of boundDoc
set seedRange to create range boundDoc start seedPoint end seedPoint
make new bookmark at boundDoc with properties {name:"wpsc_quality_splice",text object:seedRange}
set nativeRows to {{"seed",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
