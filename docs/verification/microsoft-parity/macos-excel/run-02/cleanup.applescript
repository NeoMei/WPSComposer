with timeout of 15 seconds
 tell application "/Applications/Microsoft Excel.app"
  set ownedBook to workbook "工作簿4"
  if value of range "D20" of worksheet "Data" of ownedBook is not "owned-8f8d4a8dd822" then error "Owner mismatch"
  close ownedBook saving no
  set sentinelBook to workbook "工作簿3"
  if value of range "A1" of worksheet 1 of sentinelBook is not "sentinel-8f8d4a8dd822" then error "Sentinel mismatch"
  if saved of sentinelBook then error "Sentinel unexpectedly saved"
  log "sentinel preserved; explicit fixture cleanup"
  close sentinelBook saving no
 end tell
end timeout
