use framework "Foundation"
use scripting additions
tell application "/Applications/Microsoft Excel.app"
set v to {display alerts, name of every workbook, saved of workbook "工作簿5", saved of workbook "工作簿7", value of range "A1" of worksheet 1 of workbook "工作簿5", value of range "A1" of worksheet 1 of workbook "工作簿7"}
end tell
set d to current application's NSJSONSerialization's dataWithJSONObject:v options:0 |error|:(missing value)
return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
