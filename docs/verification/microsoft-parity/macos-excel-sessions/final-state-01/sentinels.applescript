with timeout of 30 seconds
tell application "/Applications/Microsoft Excel.app"
set originalBookName to name of active workbook
set originalSheetName to name of active sheet of active workbook
set originalAddress to get address selection
set bookNames to name of every workbook
set resultText to ""
repeat with bookName in bookNames
set b to workbook (bookName as text)
if (name of b) is "工作簿5" or (name of b) is "工作簿7" then
activate object worksheet 1 of b
set resultText to resultText & (name of b) & "|" & (value of range "A1" of worksheet 1 of b) & "|" & (saved of b) & linefeed
else
set resultText to resultText & (name of b) & "|" & (full name of b) & "|" & (saved of b) & linefeed
end if
end repeat
activate object worksheet originalSheetName of workbook originalBookName
select range originalAddress of worksheet originalSheetName of workbook originalBookName
return resultText
end tell
end timeout