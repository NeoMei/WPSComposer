use framework "AppKit"
use scripting additions
with timeout of 15 seconds
 tell application "/Applications/Microsoft Excel.app"
  set wb to workbook "session-probe.xlsx"
  if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-session-probe-02rynagu/session-probe.xlsx" then error "Path mismatch"
  set ws to worksheet "Data" of wb
  set pasteboardBefore to (current application's NSPasteboard's generalPasteboard()'s changeCount()) as integer
  copy range (range "2:2" of ws) destination (range "30:30" of ws)
  if value of range "B30" of ws is not 10 then error "Copy mismatch"
  cut range (range "30:30" of ws) destination of cut (range "31:31" of ws)
  if value of range "B31" of ws is not 10 then error "Cut mismatch"
  delete range (range "31:31" of ws)
  set pasteboardAfter to (current application's NSPasteboard's generalPasteboard()'s changeCount()) as integer
  log "clipboard_change_count=" & pasteboardBefore & ":" & pasteboardAfter
  if pasteboardBefore is not pasteboardAfter then error "Clipboard changed"
  log "range_copy_cut_destination=pass"
  try
   move worksheet "Summary" of wb to before worksheet "Data" of wb
   log "sheet_move=pass:" & name of worksheet 1 of wb
  on error errText number errNum
   log "sheet_move=failed:" & errNum & ":" & errText
  end try
  return "probe=pass"
 end tell
end timeout
