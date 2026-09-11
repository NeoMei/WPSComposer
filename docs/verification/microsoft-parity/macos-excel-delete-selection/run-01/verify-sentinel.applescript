with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to workbook "工作簿28"
return ((saved of sentinelBook) as text) & (character id 31) & ((value of range "A1" of worksheet 1 of sentinelBook) as text)
end tell
end timeout
