set inputFile to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer-parity-x9mtdtv9/minimal-aea22733cf3147b2904e9d55506bb498.pptx") as alias
with timeout of 8 seconds
 tell application "Microsoft PowerPoint"
  open inputFile
  set p to presentation "minimal-aea22733cf3147b2904e9d55506bb498.pptx"
  if full name of p is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer-parity-x9mtdtv9/minimal-aea22733cf3147b2904e9d55506bb498.pptx" then error "Wrong file"
  set pageCount to count of slides of p
  if saved of p is false then error "Unexpected unsaved changes"
  close p saving no
  return pageCount
 end tell
end timeout