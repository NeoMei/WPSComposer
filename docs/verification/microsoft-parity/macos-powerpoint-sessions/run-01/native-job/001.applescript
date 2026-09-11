use framework "Foundation"
use scripting additions
on encodeJSON(itemsList)
 set jsonData to current application's NSJSONSerialization's dataWithJSONObject:itemsList options:0 |error|:(missing value)
 if jsonData is missing value then error "Native snapshot JSON encoding failed"
 return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
end encodeJSON
set inputFile to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-1n5d7z76/bound-c75828b5000b4dbab85de47d59a9ebda.pptx") as alias
with timeout of 45 seconds
 tell application "Microsoft PowerPoint"
if exists presentation "bound-c75828b5000b4dbab85de47d59a9ebda.pptx" then error "Presentation collision"
open inputFile
set ownedDoc to presentation "bound-c75828b5000b4dbab85de47d59a9ebda.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-1n5d7z76/bound-c75828b5000b4dbab85de47d59a9ebda.pptx" then error "Owned path mismatch"
return "OPEN"
 end tell
end timeout
