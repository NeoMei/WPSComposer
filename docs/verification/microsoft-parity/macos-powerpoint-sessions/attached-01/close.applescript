with timeout of 25 seconds
 tell application "Microsoft PowerPoint"
set ownedDoc to presentation "attached-8e7a6970cff44161830479cb4e906bc6.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-attach-d_8nyv30/attached-8e7a6970cff44161830479cb4e906bc6.pptx" then error "Close identity mismatch"
if not saved of ownedDoc then error "Owned presentation was not saved"
close ownedDoc saving no
return "CLOSED"
end tell
end timeout