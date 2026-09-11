with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "wpscomposer-unsaved-sentinel-4b9a9e079aa14fc291637c4d7ee1267a"
return name of sentinelBook
end tell
end timeout
