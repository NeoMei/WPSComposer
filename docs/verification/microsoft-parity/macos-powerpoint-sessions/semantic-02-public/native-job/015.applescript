use framework "Foundation"
use scripting additions
on normalizedJSON(v)
 if v is missing value then return "__NATIVE_UNAVAILABLE__"
 if class of v is list then
  set resultArray to current application's NSMutableArray's array()
  repeat with itemRef in v
   set normalizedItem to my normalizedJSON(contents of itemRef)
   resultArray's addObject:normalizedItem
  end repeat
  return resultArray
 end if
 if class of v is reference then return my normalizedJSON(contents of v)
 if class of v is integer or class of v is real or class of v is boolean or class of v is text then return v
 return v as text
end normalizedJSON
on encodeJSON(itemsList)
 set normalizedItems to my normalizedJSON(itemsList)
 set jsonData to current application's NSJSONSerialization's dataWithJSONObject:normalizedItems options:0 |error|:(missing value)
 if jsonData is missing value then error "Native snapshot JSON encoding failed"
 return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
end encodeJSON
on sessionOperation()
with timeout of 158 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-bf56e797110a4a719ae6d027e4e2ce8f.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-bf56e797110a4a719ae6d027e4e2ce8f.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-npway0fj/bound-bf56e797110a4a719ae6d027e4e2ce8f.pptx" then error "Bound presentation path changed"
set targetSlide to slide 5 of ownedDoc
set currentShape to make new shape table at end of targetSlide with properties {name:"wpscomposer-0b618076b0d341b29362edc0898dacbc",left position:40,top:260,width:450,height:130,number of rows:2,number of columns:2}
set currentCell to get cell from table object of currentShape row 1 column 1
set content of text range of text frame of shape of currentCell to "Native"
set font size of font of text range of text frame of shape of currentCell to 15
set bold of font of text range of text frame of shape of currentCell to true
set font color of font of text range of text frame of shape of currentCell to {255, 255, 255}
set fore color of fill format of shape of currentCell to {51, 68, 85}
set currentCell to get cell from table object of currentShape row 1 column 2
set content of text range of text frame of shape of currentCell to "Table"
set font size of font of text range of text frame of shape of currentCell to 15
set bold of font of text range of text frame of shape of currentCell to true
set font color of font of text range of text frame of shape of currentCell to {255, 255, 255}
set fore color of fill format of shape of currentCell to {51, 68, 85}
set currentCell to get cell from table object of currentShape row 2 column 1
set content of text range of text frame of shape of currentCell to "Value"
set font size of font of text range of text frame of shape of currentCell to 15
set currentCell to get cell from table object of currentShape row 2 column 2
set content of text range of text frame of shape of currentCell to "42"
set font size of font of text range of text frame of shape of currentCell to 15
return "CREATED"
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
