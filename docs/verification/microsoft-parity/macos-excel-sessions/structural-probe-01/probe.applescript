with timeout of 30 seconds
 tell application "/Applications/Microsoft Excel.app"
  set wb to workbook "owned-bd8b818e5d3542c3a6acb214d2ea1d49.xlsx"
  if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-74byvb0k/owned-bd8b818e5d3542c3a6acb214d2ea1d49.xlsx" then error "Identity mismatch"
  set ws to worksheet "Data" of wb
  activate object ws
  copy range (range "2:2" of ws)
  insert into range (range "4:4" of ws)
  log "row_clone=" & value of range "A4" of ws
  cut range (range "2:2" of ws)
  insert into range (range "5:5" of ws)
  log "row_move=" & value of range "A4" of ws
  try
   move worksheet "Summary" of wb to before worksheet "Data" of wb
   log "sheet_move=" & name of worksheet 1 of wb
  on error errText number errNum
   log "sheet_move_error=" & errNum & ":" & errText
  end try
  copy worksheet (worksheet "Summary" of wb) after (worksheet "Data" of wb)
  log "sheet_names=" & name of every worksheet of wb
  return "probe_completed"
 end tell
end timeout
