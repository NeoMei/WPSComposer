with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "wpscomposer-unsaved-sentinel-dde12ea594d34f0ea8cfe0fc35abd8af"
return name of sentinelBook
end tell
end timeout
