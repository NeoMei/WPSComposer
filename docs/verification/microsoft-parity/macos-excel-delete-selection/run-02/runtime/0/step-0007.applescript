use framework "Foundation"
use scripting additions
on j(v)
 if v is missing value then return "null"
 try
  set box to current application's NSArray's arrayWithObject:v
  set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
  set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
  return text 2 thru -2 of s
 on error
  try
   return my j(v as text)
  on error
   return "null"
  end try
 end try
end j
on joined(itemsList)
 set oldDelimiters to AppleScript's text item delimiters
 set AppleScript's text item delimiters to ","
 set resultText to itemsList as text
 set AppleScript's text item delimiters to oldDelimiters
 return resultText
end joined
with timeout of 60 seconds
 tell application "/Applications/Microsoft Excel.app"
set ownedBook to workbook "owned-4a23c9068e8e4fb09487ea7fb23c3912.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-acpwwea4/owned-4a23c9068e8e4fb09487ea7fb23c3912.xlsx" then error "Owned workbook identity mismatch"
set sheetJSONs to {}
set remainingCells to 100
set examined to 0
set truncated to false
repeat with sheetIndex from 1 to (count of worksheets of ownedBook)
 set ws to worksheet sheetIndex of ownedBook
 activate object ws
 set usedRng to used range of ws
 set usedRows to count of rows of usedRng
 set usedCols to count of columns of usedRng
 set firstRow to first row index of usedRng
 set firstCol to first column index of usedRng
 set cellJSONs to {}
 if true then
  repeat with r from 1 to usedRows
   repeat with c from 1 to usedCols
    if remainingCells <= 0 or examined >= 100000 then
     set truncated to true
     exit repeat
    end if
    set examined to examined + 1
    set cellValue to value of cell (firstRow + r - 1) of column (firstCol + c - 1) of ws
    if false or cellValue is not "" then
     set rng to cell (firstRow + r - 1) of column (firstCol + c - 1) of ws
set rp to properties of rng
set fp to properties of font object of rng
set rangeJSON to "{" & "\"id\":" & my j("sheet:" & sheetIndex & "/cell:" & (get address rng)) & ",\"address\":" & my j(get address rng) & ",\"value\":" & my j(value of rp) & ",\"formula\":" & my j(formula of rp) & ",\"display_text\":" & my j(string value of rp) & ",\"number_format\":" & my j(number format of rp) & ",\"horizontal_alignment\":" & my j(horizontal alignment of rp) & ",\"vertical_alignment\":" & my j(vertical alignment of rp) & ",\"wrap_text\":" & my j(wrap text of rp) & ",\"indent\":" & my j(indent level of rp) & ",\"row_height\":" & my j(row height of rp) & ",\"column_width\":" & my j(column width of rp) & ",\"merged\":" & my j(merge cells of rp) & "}"
set fontJSON to "{" & "\"name\":" & my j(name of fp) & ",\"size\":" & my j(font size of fp) & ",\"bold\":" & my j(bold of fp) & ",\"italic\":" & my j(italic of fp) & ",\"underline\":" & my j(underline of fp) & ",\"strikethrough\":" & my j(strikethrough of fp) & ",\"color\":" & my j(color of fp) & "}"
set fillJSON to "{" & "\"color\":" & my j(color of interior object of rng) & "}"
set mergeJSON to "null"
if merge cells of rp is true then set mergeJSON to my j(get address merge area of rng)
set rangeJSON to (text 1 thru -2 of rangeJSON) & ",\"font\":" & fontJSON & ",\"fill\":" & fillJSON & ",\"merge_area\":" & mergeJSON & "}"

     set end of cellJSONs to rangeJSON
     set remainingCells to remainingCells - 1
    end if
   end repeat
   if truncated then exit repeat
  end repeat
 end if
 set pp to properties of page setup object of ws
 set pageJSON to "{" & "\"top_margin\":" & my j(top margin of pp) & ",\"bottom_margin\":" & my j(bottom margin of pp) & ",\"left_margin\":" & my j(left margin of pp) & ",\"right_margin\":" & my j(right margin of pp) & ",\"orientation\":" & my j((page orientation of pp) as text) & ",\"zoom\":" & my j(zoom of pp) & ",\"fit_to_pages_wide\":" & my j(fit to pages wide of pp) & ",\"fit_to_pages_tall\":" & my j(fit to pages tall of pp) & ",\"print_area\":" & my j(print area of pp) & ",\"left_header\":" & my j(left header of pp) & ",\"center_header\":" & my j(center header of pp) & ",\"right_header\":" & my j(right header of pp) & "}"
 set chartJSONs to {}
 repeat with chartIndex from 1 to (count of chart objects of ws)
  set co to chart object chartIndex of ws
  set cp to properties of co
  set titleText to missing value
  if has title of chart of co then set titleText to chart title text of chart title of chart of co
  set chartJSON to "{" & "\"id\":" & my j("sheet:" & sheetIndex & "/chart:" & chartIndex) & ",\"index\":" & my j(chartIndex) & ",\"name\":" & my j(name of cp) & ",\"chart_type\":" & my j((chart type of chart of co) as text) & ",\"has_title\":" & my j(has title of chart of co) & ",\"title\":" & my j(titleText) & "}"
  set geometryJSON to "{" & "\"left\":" & my j(left position of cp) & ",\"top\":" & my j(top of cp) & ",\"width\":" & my j(width of cp) & ",\"height\":" & my j(height of cp) & "}"
  set end of chartJSONs to (text 1 thru -2 of chartJSON) & ",\"geometry\":" & geometryJSON & "}"
 end repeat
 set shapeJSONs to {}
 repeat with shapeIndex from 1 to (count of shapes of ws)
  set shp to shape shapeIndex of ws
  set sp to properties of shp
  set sf to properties of fill format of shp
  set sl to properties of line format of shp
  set fillJSON to "{" & "\"visible\":" & my j(visible of sf) & ",\"type\":" & my j((fill format type of sf) as text) & ",\"color\":" & my j(fore color of sf) & ",\"back_color\":" & my j(back color of sf) & ",\"transparency\":" & my j(transparency of sf) & "}"
  set lineJSON to "{" & "\"visible\":" & my j(visible of sl) & ",\"color\":" & my j(fore color of sl) & ",\"weight\":" & my j(weight of sl) & ",\"dash_style\":" & my j((dash style of sl) as text) & ",\"transparency\":" & my j(transparency of sl) & "}"
  set shapeJSON to "{" & "\"id\":" & my j("sheet:" & sheetIndex & "/shape:@name=" & (name of sp)) & ",\"index\":" & my j(shapeIndex) & ",\"shape_id\":" & my j(missing value) & ",\"name\":" & my j(name of sp) & ",\"type\":" & my j((shape type of sp) as text) & "}"
  set geometryJSON to "{" & "\"left\":" & my j(left position of sp) & ",\"top\":" & my j(top of sp) & ",\"width\":" & my j(width of sp) & ",\"height\":" & my j(height of sp) & ",\"rotation\":" & my j(rotation of sp) & ",\"z_order\":" & my j(z order position of sp) & "}"
  set end of shapeJSONs to (text 1 thru -2 of shapeJSON) & ",\"geometry\":" & geometryJSON & ",\"fill\":" & fillJSON & ",\"line\":" & lineJSON & "}"
 end repeat
 set freezeJSON to my j(freeze panes of window 1 of ownedBook)
 set sheetJSON to "{" & "\"id\":" & my j("sheet:" & sheetIndex) & ",\"index\":" & my j(sheetIndex) & ",\"name\":" & my j(name of ws) & ",\"visible\":" & my j((visible of ws) as text) & ",\"used_range\":" & my j(get address usedRng) & ",\"used_rows\":" & my j(usedRows) & ",\"used_columns\":" & my j(usedCols) & "}"
 set end of sheetJSONs to (text 1 thru -2 of sheetJSON) & ",\"page_setup\":" & pageJSON & ",\"freeze_panes\":" & freezeJSON & ",\"cells\":[" & my joined(cellJSONs) & "],\"shapes\":[" & my joined(shapeJSONs) & "],\"charts\":[" & my joined(chartJSONs) & "]}"
end repeat
set resultJSON to "{" & "\"kind\":" & my j("sheet") & ",\"name\":" & my j(name of ownedBook) & ",\"path\":" & my j("/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-excel-sessions/business-04/business.xlsx") & ",\"saved\":" & my j(saved of ownedBook) & ",\"sheet_count\":" & my j(count of worksheets of ownedBook) & ",\"cells_truncated\":" & my j(truncated) & "}"
return (text 1 thru -2 of resultJSON) & ",\"sheets\":[" & my joined(sheetJSONs) & "]}"
 end tell
end timeout
