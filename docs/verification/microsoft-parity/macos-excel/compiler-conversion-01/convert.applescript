with timeout of 30 seconds
 tell application "/Applications/Microsoft Excel.app"
  set beforeBooks to {}
  set beforeNames to name of every workbook
  repeat with beforeName in beforeNames
   set priorBook to workbook (contents of beforeName)
   set end of beforeBooks to {name of priorBook, full name of priorBook, saved of priorBook}
  end repeat
  if "generated.xlsx" is in beforeNames then error "Source workbook is already open"
  set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" with read only and editable
  if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" then error "Owned source path mismatch"
  save workbook as ownedBook filename "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/converted.pdf" file format PDF file format
  if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" then error "PDF export changed source identity"
  close ownedBook saving no
  set ownedBook to missing value
  if (count of workbooks) is not (count of beforeBooks) then error "Workbook count changed"
  repeat with beforeState in beforeBooks
   set priorBook to workbook (item 1 of beforeState)
   if full name of priorBook is not item 2 of beforeState then error "Unrelated workbook path changed"
   if saved of priorBook is not item 3 of beforeState then error "Unrelated workbook saved state changed"
  end repeat
 end tell
end timeout
return "WPSCOMPOSER_MS_OFFICE_OK:spreadsheet"
