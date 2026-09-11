with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to workbook "工作簿29"
if saved of sentinelBook then error "Sentinel saved state changed"
if (value of range "A1" of worksheet 1 of sentinelBook) is not "WPSC-DELETE-SENTINEL-db689ff8e2494398ae4d64fcd1128286" then error "Sentinel marker changed"
close sentinelBook saving no
return "closed exact sentinel"
end tell
end timeout
