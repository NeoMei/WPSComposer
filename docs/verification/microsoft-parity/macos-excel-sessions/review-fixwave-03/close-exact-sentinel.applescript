with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
if "工作簿19" is in (name of every workbook) then close workbook "工作簿19" saving no
return "closed exact sentinel"
end tell
end timeout
