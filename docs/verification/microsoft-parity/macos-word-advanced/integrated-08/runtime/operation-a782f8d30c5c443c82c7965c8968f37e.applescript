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
with timeout of 162 seconds
 tell application "Microsoft Word"
  repeat with documentIndex from 1 to (count of documents)
   set d to document documentIndex
   set end of beforeDocs to {name of d, posix full name of d, saved of d, content of text object of d}
  end repeat
  try
   open file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/advanced.docx" add to recent files false
set ownedDoc to document "advanced.docx"
if posix full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/advanced.docx" then error "Owned source path mismatch"
   repaginate ownedDoc
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/quality-69c83368567c4476a37bd6b4e29c4bbc.pdf" file format format PDF add to recent files false
set ownedDoc to missing value
repeat with documentIndex from 1 to (count of documents)
 set d to document documentIndex
 if (posix full name of d is "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/advanced.docx") or (posix full name of d is "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/quality-69c83368567c4476a37bd6b4e29c4bbc.pdf") then set ownedDoc to document (name of d)
end repeat
  on error errText number errNumber
   log "WPSC_ERROR" & tab & errNumber & tab & errText
   set failureText to "Native Word operation failed (" & errNumber & "): " & errText
  end try
  try
   if ownedDoc is not missing value then
    set closePath to posix full name of ownedDoc
    set closeName to name of ownedDoc
    if closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/quality-69c83368567c4476a37bd6b4e29c4bbc.pdf" and closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-21wxtxwe/advanced.docx" then error "Cleanup identity is outside this operation"
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
