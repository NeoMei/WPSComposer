use framework "Foundation"
use scripting additions
on encodeJSON(itemsList)
 set jsonData to current application's NSJSONSerialization's dataWithJSONObject:itemsList options:0 |error|:(missing value)
 if jsonData is missing value then error "Native snapshot JSON encoding failed"
 return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
end encodeJSON
with timeout of 45 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-c75828b5000b4dbab85de47d59a9ebda.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-c75828b5000b4dbab85de47d59a9ebda.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-1n5d7z76/bound-c75828b5000b4dbab85de47d59a9ebda.pptx" then error "Bound presentation path changed"
set snapshotRows to {{"doc", name of ownedDoc, full name of ownedDoc, saved of ownedDoc, count of slides of ownedDoc, slide width of page setup of ownedDoc}}
repeat with si from 1 to (count of slides of ownedDoc)
 set sl to slide si of ownedDoc
 set noteText to ""
 try
  set noteText to content of text range of text frame of shape 2 of notes page of sl
 end try
 set end of snapshotRows to {"slide", si, slide ID of sl, layout of sl as text, noteText, follow master background of sl, fore color of fill format of background of sl}
 repeat with sh from 1 to (count of shapes of sl)
  set shp to shape sh of sl
  set txt to ""
  set fn to ""
  set fs to 0
  set fb to false
  set fi to false
  set fu to false
  set fc to {0,0,0}
  if has text frame of shp then
   set tx to text range of text frame of shp
   set txt to content of tx
   set fn to font name of font of tx
   set fs to font size of font of tx
   set fb to bold of font of tx
   set fi to italic of font of tx
   set fu to underline of font of tx
   set fc to font color of font of tx
  end if
  set end of snapshotRows to {"shape",si,sh,name of shp,shape type of shp as text,left position of shp,top of shp,width of shp,height of shp,rotation of shp,z order position of shp,fore color of fill format of shp,visible of fill format of shp,transparency of fill format of shp,txt,fn,fs,fb,fi,fu,fc}
  if has text frame of shp then
   set tx to text range of text frame of shp
   repeat with pi from 1 to (count of paragraphs of tx)
    set end of snapshotRows to {"paragraph",si,sh,pi,content of paragraph pi of tx}
   end repeat
   set pf to paragraph format of tx
   set tf to text frame of shp
   set ln to line format of shp
   set end of snapshotRows to {"detail",si,sh,fore color of ln,weight of ln,visible of ln,transparency of ln,margin left of tf,margin right of tf,margin top of tf,margin bottom of tf,word wrap of tf,alignment of pf as text,space before of pf,space after of pf,space within of pf}
  end if
  if has table of shp then
   set nr to number of rows of shp
   set nc to number of columns of shp
   repeat with ri from 1 to nr
    repeat with ci from 1 to nc
     set cc to get cell from table object of shp row ri column ci
     set end of snapshotRows to {"cell",si,sh,nr,nc,ri,ci,content of text range of text frame of shape of cc}
    end repeat
   end repeat
  end if
 end repeat
end repeat
return my encodeJSON(snapshotRows)
 end tell
end timeout
