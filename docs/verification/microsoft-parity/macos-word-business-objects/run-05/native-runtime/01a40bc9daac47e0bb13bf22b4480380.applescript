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
set boundDoc to document "document-77a00d3c690d42fcb60cd1733d903bea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-e2k8a_fg/document-77a00d3c690d42fcb60cd1733d903bea.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set previousImageCount to count inline pictures of boundDoc
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
set insertedImage to make new inline picture at imageRange with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-e2k8a_fg/image-b384ed0c3cf64f0081733f7f1b87c666.png", link to file:false, save with document:true}
set imageIndex to count inline pictures of boundDoc
set insertedImage to inline picture imageIndex of boundDoc
set naturalWidth to width of insertedImage
set naturalHeight to height of insertedImage
set lock aspect ratio of insertedImage to true
set width of insertedImage to 84.0
set currentWidth to width of insertedImage
set currentHeight to height of insertedImage
set scaleFactor to 1.0
if scaleFactor < 1.0 then
set width of insertedImage to currentWidth * scaleFactor
set height of insertedImage to currentHeight * scaleFactor
end if
set blockRange to text object of insertedImage
set alignment of paragraph format of blockRange to align paragraph center
set keep with next of paragraph format of blockRange to true
set space after of paragraph format of blockRange to 0
set boundaryPoint to (end of content of text object of boundDoc) - 1
set followingBoundary to create range boundDoc start boundaryPoint end boundaryPoint
set content of followingBoundary to return
if imageIndex is not previousImageCount + 1 then error "WPSC_IMAGE_COUNT_DELTA_FAILED"
set nativeRows to {{"created", "inline_shape", imageIndex, previousImageCount, naturalWidth, naturalHeight, width of insertedImage, height of insertedImage, alternative text of insertedImage as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
