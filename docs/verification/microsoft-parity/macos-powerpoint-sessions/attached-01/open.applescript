set inputFile to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-attach-d_8nyv30/attached-8e7a6970cff44161830479cb4e906bc6.pptx") as alias
with timeout of 25 seconds
 tell application "Microsoft PowerPoint"
open inputFile
if full name of active presentation is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-attach-d_8nyv30/attached-8e7a6970cff44161830479cb4e906bc6.pptx" then error "Active identity mismatch"
return "OPEN"
end tell
end timeout