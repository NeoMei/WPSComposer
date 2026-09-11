with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to workbook "工作簿30"
if saved of sentinelBook then error "Sentinel saved state changed"
if (value of range "A1" of worksheet 1 of sentinelBook) is not "WPSC-DELETE-SENTINEL-e639c1eb80444a91b9f8df73938f63d7" then error "Sentinel marker changed"
close sentinelBook saving no
return "closed exact sentinel"
end tell
end timeout
