use framework "Foundation"
use scripting additions
tell application "/Applications/Microsoft Excel.app"
set names to name of every workbook
end tell
set payload to current application's NSArray's arrayWithArray:names
set jsonData to current application's NSJSONSerialization's dataWithJSONObject:payload options:0 |error|:(missing value)
return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
