with timeout of 40 seconds
 tell application "/Applications/Microsoft Excel.app"
  set beforeBooks to {}
  set beforeNames to name of every workbook
  repeat with beforeName in beforeNames
   set priorBook to workbook (contents of beforeName)
   set end of beforeBooks to {name of priorBook, full name of priorBook, saved of priorBook}
  end repeat
  if "generated.xlsx" is in beforeNames then error "Target workbook name is already open"
  set ownedBook to make new workbook
  set ownedName to name of ownedBook
  if ownedName is in beforeNames then error "New workbook identity collision"
  set ownedBook to workbook ownedName
  repeat while (count of worksheets of ownedBook) > 1
   delete worksheet 2 of ownedBook
  end repeat
  set name of worksheet 1 of ownedBook to "Sheet1"
  set ownedSheet to worksheet 1 of ownedBook
  set name of worksheet 1 of ownedBook to "Data"
  set ownedSheet to worksheet 1 of ownedBook
  set value of range "B2:C4" of ownedSheet to {{"Metric", "Value"}, {"Alpha", 10}, {"Total", "=SUM(C3:C3)"}}
  set font size of font object of range "B2:C4" of ownedSheet to 12
  set bold of font object of range "B2:C2" of ownedSheet to true
  set color of font object of range "B2:C2" of ownedSheet to {255, 255, 255}
  set color of interior object of range "B2:C2" of ownedSheet to {68, 114, 196}
  set column width of range "B:C" of ownedSheet to 22
  autofit entire column of used range of ownedSheet
  tell ownedBook
   make new worksheet at end with properties {name:"Summary"}
  end tell
  set ownedSheet to worksheet 2 of ownedBook
  set value of range "A1:B2" of ownedSheet to {{"Cross sheet total", "Value"}, {"Double", "=Data!C4*2"}}
  set font size of font object of range "A1:B2" of ownedSheet to 11
  set bold of font object of range "A1:B1" of ownedSheet to true
  set color of font object of range "A1:B1" of ownedSheet to {255, 255, 255}
  autofit entire column of used range of ownedSheet
  set ownedSheet to worksheet 1 of ownedBook
  repeat with sheetIndex from 1 to (count of worksheets of ownedBook)
   calculate worksheet sheetIndex of ownedBook
  end repeat
  save workbook as ownedBook filename "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" file format Excel XML file format
  set ownedBook to workbook "generated.xlsx"
  if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" then error "Owned saved path mismatch"
  save workbook as ownedBook filename "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.pdf" file format PDF file format
  if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-compiler-_oxf6dcq/generated.xlsx" then error "PDF export changed workbook identity"
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
