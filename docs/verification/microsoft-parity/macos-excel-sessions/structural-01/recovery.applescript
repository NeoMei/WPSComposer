with timeout of 20 seconds
 tell application "/Applications/Microsoft Excel.app"
  set wb to workbook "owned-59f177c266d247248f5ff1880eb76f75.xlsx"
  if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-mtndu0cx/owned-59f177c266d247248f5ff1880eb76f75.xlsx" then error "Recovery path mismatch"
  save wb
  close wb saving no
  return "explicit_owned_recovery_closed"
 end tell
end timeout
