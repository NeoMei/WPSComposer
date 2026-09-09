with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
close workbook "工作簿28" saving no
return "closed exact sentinel"
end tell
end timeout
