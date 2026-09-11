with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "WPSC-DELETE-SENTINEL-db689ff8e2494398ae4d64fcd1128286"
return name of sentinelBook
end tell
end timeout
