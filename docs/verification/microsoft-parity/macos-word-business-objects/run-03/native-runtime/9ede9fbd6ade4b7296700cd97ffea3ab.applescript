use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex
with timeout of 60 seconds
tell application "/Applications/Microsoft Word.app"
set nativeRows to {}
set boundDoc to document "document-ad7a443c7ba54d44a51f6dbf1f981020.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-i9e61gi1/document-ad7a443c7ba54d44a51f6dbf1f981020.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set imageRange to create range boundDoc start insertionPoint end insertionPoint
set insertedImage to make new picture at boundDoc with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-i9e61gi1/image-555aea008450462abf5e7bbe404d09f9.png", link to file:false, save with document:true, anchor:imageRange}
set imageIndex to count shapes of boundDoc
set insertedImage to shape imageIndex of boundDoc
set relative horizontal position of insertedImage to relative horizontal position page
set relative vertical position of insertedImage to relative vertical position page
set left position of insertedImage to 0
set top of insertedImage to 0
set wrap type of wrap format of insertedImage to wrap square
set naturalWidth to width of insertedImage
set naturalHeight to height of insertedImage
set lock aspect ratio of insertedImage to false
set width of insertedImage to 96.0
set height of insertedImage to 54.0
set currentWidth to width of insertedImage
set currentHeight to height of insertedImage
set scaleFactor to 1.0
if scaleFactor < 1.0 then
set width of insertedImage to currentWidth * scaleFactor
set height of insertedImage to currentHeight * scaleFactor
end if
set nativeRows to {{"created", "shape", imageIndex, width of insertedImage, height of insertedImage, my enumIndex(wrap type of wrap format of insertedImage, {wrap square, wrap tight, wrap through, wrap none, wrap top bottom, wrap behind, wrap front, wrap inline})}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
