with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set b to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-active-fixture-2c8jri01/active-owned.xlsx"
activate object worksheet "Data" of b
select range "B2" of worksheet "Data" of b
return full name of b
end tell
end timeout