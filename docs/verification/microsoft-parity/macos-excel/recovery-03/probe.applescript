with timeout of 40 seconds
 tell application "/Applications/Microsoft Excel.app"
  set targetPath to "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-03/native.xlsx"
  set ownedBook to workbook "native.xlsx"
  if full name of ownedBook is not targetPath then error "Owned path mismatch"
  if value of range "D20" of worksheet "Data" of ownedBook is not "owned-58d9ceb7df8e" then error "Owned content mismatch"
  set sentinelBook to workbook "工作簿5"
  if value of range "A1" of worksheet 1 of sentinelBook is not "sentinel-58d9ceb7df8e" then error "Sentinel mismatch"
  if saved of sentinelBook then error "Sentinel saved unexpectedly"
  close ownedBook saving no
  log "owned_close=pass"
  set ownedBook to open workbook workbook file name targetPath with read only
  if full name of ownedBook is not targetPath then error "Reopen path mismatch"
  set dataSheet to worksheet "Data" of ownedBook
  set summarySheet to worksheet "Summary" of ownedBook
  if value of range "D20" of dataSheet is not "owned-58d9ceb7df8e" then error "Reopen owner mismatch"
  if formula of range "B6" of dataSheet is not "=SUM(B2:B4)" then error "Reopen formula mismatch"
  if value of range "B6" of dataSheet is not 60 then error "Reopen calculation mismatch"
  if value of range "B1" of summarySheet is not 120 then error "Cross sheet mismatch"
  if merge cells of range "A8:D8" of dataSheet is not true then error "Merge mismatch"
  if count of chart objects of dataSheet is not 1 then error "Native chart missing"
  log "reopen_semantics=pass"
  save workbook as ownedBook filename "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-03/recovered.pdf" file format PDF file format
  log "pdf_export_returned=pass"
  if value of range "A1" of worksheet 1 of sentinelBook is not "sentinel-58d9ceb7df8e" then error "Sentinel changed"
  if saved of sentinelBook then error "Sentinel saved unexpectedly"
  log "sentinel_preserved=pass"
  close ownedBook saving no
  return "recovery=pass"
 end tell
end timeout
