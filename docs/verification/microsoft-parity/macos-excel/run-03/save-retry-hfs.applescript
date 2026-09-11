with timeout of 12 seconds
 tell application "/Applications/Microsoft Excel.app"
  set ownedBook to workbook "native.xlsx"
  if value of range "D20" of worksheet "Data" of ownedBook is not "owned-58d9ceb7df8e" then error "Owner mismatch"
  set destinationPath to (POSIX file "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel/run-03/native-hfs.xlsx") as text
  log destinationPath
  save workbook as ownedBook filename destinationPath file format Excel XML file format
  return "save_hfs=pass"
 end tell
end timeout
