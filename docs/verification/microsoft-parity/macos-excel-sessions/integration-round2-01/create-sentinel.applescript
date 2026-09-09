with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "integration-round2-sentinel-23acd7d1e8e8430e9c91894ed3f09b16"
return name of sentinelBook
end tell
end timeout
