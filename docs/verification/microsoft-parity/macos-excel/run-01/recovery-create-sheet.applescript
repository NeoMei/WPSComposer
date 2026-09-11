with timeout of 15 seconds
 tell application "/Applications/Microsoft Excel.app"
  set ownedBook to workbook "工作簿2"
  if value of range "D20" of worksheet "Data" of ownedBook is not "owned-95ef52181d1b" then error "Owner mismatch"
  tell ownedBook
   set ws to make new worksheet at end with properties {name:"Summary"}
  end tell
  return name of ws
 end tell
end timeout
