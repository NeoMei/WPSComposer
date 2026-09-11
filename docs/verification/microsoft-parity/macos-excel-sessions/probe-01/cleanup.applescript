with timeout of 10 seconds
 tell application "/Applications/Microsoft Excel.app"
  set wb to workbook "session-probe.xlsx"
  if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-session-probe-02rynagu/session-probe.xlsx" then error "Path mismatch"
  if value of range "D20" of worksheet "Data" of wb is not "owned-e9a7d3dce1ad" then error "Content mismatch"
  close wb saving no
  return "owned_probe_closed"
 end tell
end timeout
