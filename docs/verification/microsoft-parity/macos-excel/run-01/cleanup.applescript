with timeout of 15 seconds
 tell application "/Applications/Microsoft Excel.app"
  set ownedBook to workbook "工作簿2"
  if value of range "D20" of worksheet "Data" of ownedBook is not "owned-95ef52181d1b" then error "Owner mismatch"
  close ownedBook saving no
  set sentinelBook to workbook "工作簿1"
  if value of range "A1" of worksheet 1 of sentinelBook is not "sentinel-95ef52181d1b" then error "Sentinel mismatch"
  if saved of sentinelBook then error "Sentinel unexpectedly saved"
  log "sentinel preserved; explicit fixture cleanup"
  close sentinelBook saving no
 end tell
end timeout
