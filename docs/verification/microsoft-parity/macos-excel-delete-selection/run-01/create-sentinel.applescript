with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "WPSC-DELETE-SENTINEL-36a9e0ddd3024a57ad8cdc5290e26336"
return name of sentinelBook
end tell
end timeout
