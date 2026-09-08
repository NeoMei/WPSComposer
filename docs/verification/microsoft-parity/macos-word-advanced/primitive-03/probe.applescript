on appendText(ownedDoc, valueText)
 tell application "Microsoft Word"
  set p to (end of content of text object of ownedDoc) - 1
  set r to create range ownedDoc start p end p
  set content of r to valueText
  set e to (end of content of text object of ownedDoc) - 1
  return create range ownedDoc start p end e
 end tell
end appendText
set ownedDoc to missing value
set beforeDocs to {}
set failureText to ""
with timeout of 60 seconds
 tell application "Microsoft Word"
  repeat with documentIndex from 1 to (count of documents)
   set d to document documentIndex
   set end of beforeDocs to {name of d, posix full name of d, saved of d, content of text object of d}
  end repeat
  try
   set ownedDoc to make new document
set ownedName to name of ownedDoc
repeat with prior in beforeDocs
 if item 1 of prior is ownedName then
  set ownedDoc to missing value
  error "New document identity collision"
 end if
end repeat
set ownedDoc to document ownedName
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-jvs917fb/primitive.docx" file format format document default add to recent files false
set ownedDoc to document "primitive.docx"
if posix full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-jvs917fb/primitive.docx" then error "Owned document path mismatch"
   set r to my appendText(ownedDoc, "FORMATTED-CONTENT" & return)
set color of font object of r to {17, 85, 153}
set underline of font object of r to underline single
set strike through of font object of r to true
set (left indent of paragraph format of r) to 18
set right indent of paragraph format of r to 9
set line spacing rule of paragraph format of r to line space exactly
set line spacing of paragraph format of r to 20
set background pattern color of shading of paragraph format of r to {230, 240, 250}
set ownBorder to get border (paragraph format of r) which border border left
set line style of ownBorder to line style single
set color of ownBorder to {17, 85, 153}
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
set ownTable to make new table at ownedDoc with properties {text object:r, number of rows:3, number of columns:3}
set c1 to get cell from table ownTable row 2 column 1
set c2 to get cell from table ownTable row 2 column 2
set content of text object of c1 to "MERGED-CONTENT"
merge cell c1 with c2
set r to my appendText(ownedDoc, return & "AFTER-MERGE" & return)
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
insert break at r break type section break next page
set orientation of page setup of section 2 of ownedDoc to orient landscape
set number of text columns (page setup of section 2 of ownedDoc) number of columns 2
set r to my appendText(ownedDoc, "COLUMN-ONE" & return)
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
insert break at r break type column break
set r to my appendText(ownedDoc, "COLUMN-TWO" & return)
save ownedDoc
  on error errText number errNumber
   log "WPSC_ERROR" & tab & errNumber & tab & errText
   set failureText to "Native Word operation failed (" & errNumber & "): " & errText
  end try
  try
   if ownedDoc is not missing value then
    set closePath to posix full name of ownedDoc
    set closeName to name of ownedDoc
    if closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-jvs917fb/primitive.docx" and closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-jvs917fb/primitive.docx" then error "Cleanup identity is outside this operation"
    close document closeName saving no
   end if
   set ownedDoc to missing value
  on error
   error "Native Word owned-document cleanup failed; staging quarantined"
  end try
  if (count of documents) is not (count of beforeDocs) then error "Native Word document count changed; staging quarantined"
  repeat with prior in beforeDocs
   set d to document (item 1 of prior)
   if posix full name of d is not item 2 of prior then error "Preexisting Word path changed"
   if saved of d is not item 3 of prior then error "Preexisting Word saved state changed"
   if content of text object of d is not item 4 of prior then error "Preexisting Word text changed"
  end repeat
  log "WPSC_CLEAN"
  if failureText is not "" then error failureText
 end tell
end timeout
return "WPSC_OK" & tab & (count of beforeDocs)
