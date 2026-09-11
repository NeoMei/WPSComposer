use framework "Foundation"
use scripting additions
on j(v)
 if v is missing value then return "null"
 set box to current application's NSArray's arrayWithObject:v
 set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
 set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
 return text 2 thru -2 of s
end j
with timeout of 30 seconds
 tell application "/Applications/Microsoft Excel.app"
  set wb to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-session-probe-02rynagu/session-probe.xlsx" with editable
  if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/Documents/wpscomposer-session-probe-02rynagu/session-probe.xlsx" then error "Path mismatch"
  set ws to worksheet "Data" of wb
  set r to range "A1" of ws
  set rp to properties of r
  log my j(value of rp)
  log my j(properties of font object of r)
  log "underline=" & ((underline of font object of r) as text)
  log "alignment=" & ((horizontal alignment of r) as text)
  log "chart=" & ((chart type of chart of chart object 1 of ws) as text)
  log "used=" & (get address used range of ws)
  return "probe=pass"
 end tell
end timeout
