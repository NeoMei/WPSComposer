with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "integration-round2-sentinel-eadad49bbdc2474e9da177aaa2d74dba"
return name of sentinelBook
end tell
end timeout
