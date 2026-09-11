tell application "Microsoft Word"
set ownedDoc to document "document-2b532297486e4b5789d4a83f9e6f8ad4.docx"
if (posix full name of ownedDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jeypd1s4/document-2b532297486e4b5789d4a83f9e6f8ad4.docx" then error "IDENTITY_CHANGED"
if not saved of ownedDoc then error "UNSAVED_CHANGED"
close ownedDoc saving no
end tell
return "EXACT_READONLY_COPY_CLOSED"