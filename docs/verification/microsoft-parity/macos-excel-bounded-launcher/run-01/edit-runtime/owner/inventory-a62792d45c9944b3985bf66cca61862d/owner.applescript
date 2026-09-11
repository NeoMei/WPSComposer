use framework "Foundation"
use scripting additions
on j(v)
 set box to current application's NSArray's arrayWithObject:v
 set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
 set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
 return text 2 thru -2 of s
end j
on pristine(b)
 tell application "/Applications/Microsoft Excel.app"
  try
   if saved of b is not true or path of b is not "" or full name of b is not name of b then return false
   if has vb project of b or is add in of b then return false
   if (count of sheets of b) is not 1 or (count of worksheets of b) is not 1 then return false
   if (count of named items of b) is not 0 then return false
   set s to worksheet 1 of b
   if (count of shapes of s) is not 0 or (count of chart objects of s) is not 0 then return false
   if (count of named items of s) is not 0 or (count of Excel comments of s) is not 0 then return false
   if (count of hyperlinks of s) is not 0 or (count of query tables of s) is not 0 then return false
   if (count of cells of used range of s) is not 1 then return false
   if (get address of used range of s) is not "$A$1" then return false
   if value of range "A1" of s is not "" or formula of range "A1" of s is not "" then return false
   return true
  on error
   return false
  end try
 end tell
end pristine

-- owner operation: inventory
set ownerNonce to "a62792d45c9944b3985bf66cca61862d"
tell application "/Applications/Microsoft Excel.app"
set rowsJSON to ""
repeat with ownerIndex from 1 to (count of workbooks)
 set b to workbook ownerIndex
 if rowsJSON is not "" then set rowsJSON to rowsJSON & ","
 set rowsJSON to rowsJSON & "{\"name\":" & my j(name of b) & ",\"path\":" & my j(path of b) & ",\"full_name\":" & my j(full name of b) & ",\"saved\":" & my j(saved of b) & ",\"pristine\":" & my j(my pristine(b)) & "}"
end repeat
return "{\"nonce\":" & my j(ownerNonce) & ",\"version\":" & my j(version as text) & ",\"workbooks\":[" & rowsJSON & "]}"
end tell
