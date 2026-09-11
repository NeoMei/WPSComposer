with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to "WPSC-DELETE-SENTINEL-e639c1eb80444a91b9f8df73938f63d7"
return name of sentinelBook
end tell
end timeout
