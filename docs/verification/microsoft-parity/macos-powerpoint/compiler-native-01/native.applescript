set nativePath to "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer-compiler-32wf7tvf/compiler-770f7573f9c0409d8ee46cc88fdd4e16.pptx"
set nativeHFS to (POSIX file nativePath) as text
set ownedDoc to missing value
set pdfPath to "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer-compiler-32wf7tvf/native.pdf"
set pdfHFS to (POSIX file pdfPath) as text
with timeout of 30 seconds
 tell application "Microsoft PowerPoint"
  set beforeDocs to {}
repeat with snapshotIndex from 1 to (count of presentations)
 set end of beforeDocs to {name of presentation snapshotIndex, full name of presentation snapshotIndex, saved of presentation snapshotIndex, count of slides of presentation snapshotIndex}
end repeat
repeat with prior in beforeDocs
 if item 1 of prior is "compiler-770f7573f9c0409d8ee46cc88fdd4e16.pptx" then error "Existing presentation collision"
end repeat
  try
   set ownedDoc to make new presentation
set createdName to name of ownedDoc
repeat with prior in beforeDocs
 if item 1 of prior is createdName then error "New presentation collision"
end repeat
   repeat while (count of slides of ownedDoc) > 0
 delete slide 1 of ownedDoc
end repeat
set slide size of page setup of ownedDoc to slide size on screen
set slide width of page setup of ownedDoc to 960
set fore color of fill format of background of slide master of ownedDoc to {255, 255, 255}
set currentSlide to make new slide at end of ownedDoc with properties {layout:slide layout title slide}
set follow master background of currentSlide to false
set fore color of fill format of background of currentSlide to {255, 255, 255}
set content of text range of text frame of shape 1 of currentSlide to "Compiler native title"
set font size of font of text range of text frame of shape 1 of currentSlide to 36
set font color of font of text range of text frame of shape 1 of currentSlide to {0, 82, 148}
set bold of font of text range of text frame of shape 1 of currentSlide to false
set font name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set east asian name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set content of text range of text frame of shape 2 of currentSlide to "Chinese 中文 subtitle"
set font size of font of text range of text frame of shape 2 of currentSlide to 20
set font color of font of text range of text frame of shape 2 of currentSlide to {90, 90, 90}
set bold of font of text range of text frame of shape 2 of currentSlide to false
set font name of font of text range of text frame of shape 2 of currentSlide to "Arial"
set east asian name of font of text range of text frame of shape 2 of currentSlide to "Arial"
set currentSlide to make new slide at end of ownedDoc with properties {layout:slide layout title only}
set follow master background of currentSlide to false
set fore color of fill format of background of currentSlide to {255, 255, 255}
set content of text range of text frame of shape 1 of currentSlide to "Section title"
set font size of font of text range of text frame of shape 1 of currentSlide to 36
set font color of font of text range of text frame of shape 1 of currentSlide to {0, 82, 148}
set bold of font of text range of text frame of shape 1 of currentSlide to false
set font name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set east asian name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set currentSlide to make new slide at end of ownedDoc with properties {layout:slide layout text slide}
set follow master background of currentSlide to false
set fore color of fill format of background of currentSlide to {255, 255, 255}
set content of text range of text frame of shape 1 of currentSlide to "Native bullets"
set font size of font of text range of text frame of shape 1 of currentSlide to 36
set font color of font of text range of text frame of shape 1 of currentSlide to {0, 82, 148}
set bold of font of text range of text frame of shape 1 of currentSlide to false
set font name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set east asian name of font of text range of text frame of shape 1 of currentSlide to "Arial"
set content of text range of text frame of shape 2 of currentSlide to "Alpha" & return & "Beta 中文"
set font size of font of text range of text frame of shape 2 of currentSlide to 18
set font color of font of text range of text frame of shape 2 of currentSlide to {45, 45, 48}
set bold of font of text range of text frame of shape 2 of currentSlide to false
set font name of font of text range of text frame of shape 2 of currentSlide to "Arial"
set east asian name of font of text range of text frame of shape 2 of currentSlide to "Arial"
set currentSlide to make new slide at end of ownedDoc with properties {layout:slide layout blank}
set follow master background of currentSlide to false
set fore color of fill format of background of currentSlide to {255, 255, 255}
set targetSlide to slide 4 of ownedDoc
set currentPicture to make new picture at end of targetSlide with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer-compiler-32wf7tvf/image.png", link to file:false, save with document:true, left position:40, top:40}
set lock aspect ratio of currentPicture to true
set width of currentPicture to 100
set targetSlide to slide 4 of ownedDoc
set currentTable to make new shape table at end of targetSlide with properties {number of rows:2, number of columns:2, left position:200, top:80, width:450, height:180}
set currentCell to get cell from table object of currentTable row 1 column 1
set content of text range of text frame of shape of currentCell to "Name"
set font size of font of text range of text frame of shape of currentCell to 16
set font color of font of text range of text frame of shape of currentCell to {255, 255, 255}
set bold of font of text range of text frame of shape of currentCell to true
set fore color of fill format of shape of currentCell to {68, 114, 196}
set currentCell to get cell from table object of currentTable row 1 column 2
set content of text range of text frame of shape of currentCell to "Value"
set font size of font of text range of text frame of shape of currentCell to 16
set font color of font of text range of text frame of shape of currentCell to {255, 255, 255}
set bold of font of text range of text frame of shape of currentCell to true
set fore color of fill format of shape of currentCell to {68, 114, 196}
set currentCell to get cell from table object of currentTable row 2 column 1
set content of text range of text frame of shape of currentCell to "Native"
set font size of font of text range of text frame of shape of currentCell to 16
set font color of font of text range of text frame of shape of currentCell to {0, 0, 0}
set bold of font of text range of text frame of shape of currentCell to false
set currentCell to get cell from table object of currentTable row 2 column 2
set content of text range of text frame of shape of currentCell to "42"
set font size of font of text range of text frame of shape of currentCell to 16
set font color of font of text range of text frame of shape of currentCell to {0, 0, 0}
set bold of font of text range of text frame of shape of currentCell to false
   save ownedDoc in nativeHFS as save as Open XML presentation
set ownedDoc to presentation "compiler-770f7573f9c0409d8ee46cc88fdd4e16.pptx"
if full name of ownedDoc is not nativePath then error "Owned output path mismatch"
if saved of ownedDoc is false then error "Owned presentation was not saved"
save ownedDoc in pdfHFS as save as PDF

   close ownedDoc saving no
   set ownedDoc to missing value
   if (count of presentations) is not (count of beforeDocs) then error "Unrelated presentation count changed"
repeat with snapshotIndex from 1 to (count of beforeDocs)
 if {name of presentation snapshotIndex, full name of presentation snapshotIndex, saved of presentation snapshotIndex, count of slides of presentation snapshotIndex} is not item snapshotIndex of beforeDocs then error "Unrelated presentation state changed"
end repeat
   return "WPSCOMPOSER_MS_OFFICE_OK:presentation"
  on error errorText number errorNumber
   log "WPSCOMPOSER_MS_OFFICE_ERROR:presentation" & tab & errorNumber & tab & errorText
   error errorText number errorNumber
  end try
 end tell
end timeout
