with timeout of 15 seconds
 tell application "Microsoft Word"
  set sentinelDoc to make new document
  set initialText to content of text object of sentinelDoc
  set initialName to name of sentinelDoc
  set initialPath to posix full name of sentinelDoc
  set initialSaved to saved of sentinelDoc
  try
   set content of text object of sentinelDoc to "WPSC-ADVANCED-SENTINEL-9479832e36bd45a3ba0d341763b9b776"
   set sentinelText to content of text object of sentinelDoc
   return name of sentinelDoc & tab & posix full name of sentinelDoc & tab & (saved of sentinelDoc as text)
  on error errText number errNumber
   if name of sentinelDoc is initialName and posix full name of sentinelDoc is initialPath and saved of sentinelDoc is initialSaved and content of text object of sentinelDoc is initialText then close sentinelDoc saving no
   error errText number errNumber
  end try
 end tell
end timeout