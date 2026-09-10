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

-- owner operation: close
set ownerNonce to "1872e3174b8e4b508e3bdd044bc30a65"
tell application "/Applications/Microsoft Excel.app"
if (count of workbooks) is not 1 then error "Foreign workbooks"
set b to workbook "owned-3e33878580de48ba82e8c5758298d6d8.xlsx"
if full name of b is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-kisnvreg/owned-3e33878580de48ba82e8c5758298d6d8.xlsx" then error "Workbook identity mismatch"
close b saving no
return "{\"nonce\":" & my j(ownerNonce) & ",\"empty\":" & my j((count of workbooks) is 0) & "}"
end tell
