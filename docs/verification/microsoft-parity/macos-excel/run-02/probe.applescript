with timeout of 58 seconds
 tell application "/Applications/Microsoft Excel.app"
  set previousBooks to {}
  repeat with existingBook in workbooks
   set end of previousBooks to {name of existingBook, full name of existingBook, saved of existingBook}
  end repeat
  log "before_books=" & (count of previousBooks)
  set sentinelToken to "sentinel-8f8d4a8dd822"
  set ownedToken to "owned-8f8d4a8dd822"
  set sentinelBook to make new workbook
  set sentinelName to name of sentinelBook
  set value of range "A1" of worksheet 1 of sentinelBook to sentinelToken
  log "sentinel_name=" & sentinelName
  set sentinelPath to full name of sentinelBook
  set ownedBook to make new workbook
  set ownedName to name of ownedBook
  set dataSheet to worksheet 1 of ownedBook
  set name of dataSheet to "Data"
  set value of range "D20" of dataSheet to ownedToken
  log "owned_name=" & ownedName
  log "stage=data"
  set value of range "A1:B4" of dataSheet to {{"Month", "Revenue"}, {"Jan", 10}, {"Feb", 20}, {"Mar", 30}}
  set value of range "A6" of dataSheet to "Total"
  set formula of range "B6" of dataSheet to "=SUM(B2:B4)"
  tell ownedBook
   make new worksheet at end with properties {name:"Summary"}
  end tell
  set summarySheet to worksheet "Summary" of ownedBook
  set value of range "A1" of summarySheet to "Cross sheet total"
  set formula of range "B1" of summarySheet to "=Data!B6*2"
  set value of range "A8" of dataSheet to "Microsoft native parity"
  merge range "A8:D8" of dataSheet
  set bold of font object of range "A1:B1" of dataSheet to true
  set color of interior object of range "A1:B1" of dataSheet to {220, 235, 250}
  set number format of range "B2:B6" of dataSheet to "0.00"
  set column width of range "A:D" of dataSheet to 18
  set row height of range "1:1" of dataSheet to 28
  calculate dataSheet
  calculate summarySheet
  if value of range "B6" of dataSheet is not 60 then error "Formula value mismatch"
  if value of range "B1" of summarySheet is not 120 then error "Cross sheet value mismatch"
  log "formula_calculation=pass"
  log "stage=chart"
  set chartContainer to make new chart object at end of chart objects of dataSheet with properties {left position:20, top:160, width:400, height:220}
  set nativeChart to chart of chartContainer
  set chart type of nativeChart to column clustered
  set source data nativeChart source range "A1:B4" of dataSheet plot by columns
  set has title of nativeChart to true
  set chart title text of chart title of nativeChart to "Revenue by month"
  log "chart=pass"
  log "stage=structure"
  try
   insert into range (range "10:10" of dataSheet)
   set value of range "A10" of dataSheet to "Inserted row"
   if value of range "D21" of dataSheet is not ownedToken then error "Row insertion lost marker"
   delete range (range "10:10" of dataSheet)
   if value of range "D20" of dataSheet is not ownedToken then error "Row deletion lost marker"
   log "row_insert_delete=pass"
  on error messageText number errorNumber
   log "row_insert_delete=failed:" & errorNumber & ":" & messageText
  end try
  try
   tell ownedBook
    make new worksheet at end with properties {name:"Temporary"}
   end tell
   set temporarySheet to worksheet "Temporary" of ownedBook
   set name of temporarySheet to "Renamed"
   copy worksheet temporarySheet after temporarySheet
   if count of worksheets of ownedBook is not 4 then error "Sheet clone count mismatch"
   delete worksheet "Renamed (2)" of ownedBook
   delete worksheet "Renamed" of ownedBook
   if count of worksheets of ownedBook is not 2 then error "Sheet removal count mismatch"
   log "sheet_create_rename_clone_delete=pass"
  on error messageText number errorNumber
   log "sheet_create_rename_clone_delete=failed:" & errorNumber & ":" & messageText
  end try
  try
   activate object dataSheet
   select range "A1:B4" of dataSheet
   set selectionAddress to get address selection
   if selectionAddress is not "$A$1:$B$4" then error "Selection mismatch"
   log "selection=pass:" & selectionAddress
  on error messageText number errorNumber
   log "selection=failed:" & errorNumber & ":" & messageText
  end try
  set print area of page setup object of dataSheet to "$A$1:$F$28"
  set zoom of page setup object of dataSheet to false
  set fit to pages wide of page setup object of dataSheet to 1
  set fit to pages tall of page setup object of dataSheet to 1
  set print area of page setup object of summarySheet to "$A$1:$D$4"
  log "stage=save"
  if value of range "D20" of dataSheet is not ownedToken then error "Owned marker missing before save"
  save workbook as ownedBook filename "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-02/native.xlsx" file format Excel XML file format
  set ownedBook to workbook "native.xlsx"
  if full name of ownedBook is not "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-02/native.xlsx" then error "Saved path identity mismatch"
  close ownedBook saving no
  log "stage=reopen"
  set ownedBook to open workbook workbook file name "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-02/native.xlsx" with read only
  if full name of ownedBook is not "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-02/native.xlsx" then error "Reopen path identity mismatch"
  set dataSheet to worksheet "Data" of ownedBook
  set summarySheet to worksheet "Summary" of ownedBook
  if value of range "D20" of dataSheet is not ownedToken then error "Reopen owner marker missing"
  if formula of range "B6" of dataSheet is not "=SUM(B2:B4)" then error "Reopen formula mismatch"
  if value of range "B6" of dataSheet is not 60 then error "Reopen calculated value mismatch"
  if value of range "B1" of summarySheet is not 120 then error "Reopen cross sheet value mismatch"
  if merge cells of range "A8:D8" of dataSheet is not true then error "Reopen merged range mismatch"
  if bold of font object of range "A1:B1" of dataSheet is not true then error "Reopen font mismatch"
  if count of chart objects of dataSheet is not 1 then error "Reopen native chart missing"
  log "reopen_semantics=pass"
  log "stage=pdf"
  save workbook as ownedBook filename "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-02/native.pdf" file format PDF file format
  log "pdf_export=pass"
  if value of range "A1" of worksheet 1 of sentinelBook is not sentinelToken then error "Sentinel content changed"
  if saved of sentinelBook is not false then error "Sentinel was saved"
  if full name of sentinelBook is not sentinelPath then error "Sentinel path changed"
  repeat with originalState in previousBooks
   set originalBook to workbook (item 1 of originalState)
   if full name of originalBook is not item 2 of originalState then error "Unrelated path changed"
   if saved of originalBook is not item 3 of originalState then error "Unrelated saved state changed"
  end repeat
  log "unrelated_and_unsaved_sentinel=pass"
  if value of range "D20" of worksheet "Data" of ownedBook is not ownedToken then error "Owned marker missing before close"
  close ownedBook saving no
  log "owned_closed=pass"
  -- Sentinel intentionally remains open and unsaved as preservation evidence.
  return "native_probe=pass"
 end tell
end timeout
