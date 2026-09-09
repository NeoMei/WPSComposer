with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set inventoryRows to {}
repeat with workbookIndex from 1 to (count of workbooks)
 set candidateBook to workbook workbookIndex
 try
  set candidatePath to full name of candidateBook
 on error
  set candidatePath to "__UNSAVED__"
 end try
 set end of inventoryRows to ((name of candidateBook) & (character id 31) & ((saved of candidateBook) as text) & (character id 31) & candidatePath)
end repeat
set AppleScript's text item delimiters to character id 30
return inventoryRows as text
end tell
end timeout
