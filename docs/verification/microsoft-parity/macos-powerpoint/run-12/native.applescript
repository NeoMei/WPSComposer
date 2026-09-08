set pptAlias to POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-58ikqxox/native-f39221af45564258bb2b1c76221b2eed.pptx"
set pptPath to pptAlias as text
set pdfPath to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-58ikqxox/native.pdf") as text
with timeout of 35 seconds
 tell application "Microsoft PowerPoint"
  set beforeState to {}
  repeat with existing in (get presentations)
   set end of beforeState to {name of existing, full name of existing, saved of existing, count of slides of existing}
  end repeat
  log "PROBE|baseline|" & (count of beforeState)
  log "PROBE|version|" & Version
  set sentinel to make new presentation
  set sentSlide to make new slide at end of sentinel with properties {layout:slide layout blank}
  set sentText to make new text box at end of sentSlide with properties {left position:40, top:40, width:500, height:80}
  set content of text range of text frame of sentText to "SENTINEL-f39221af45564258bb2b1c76221b2eed"
  set sentinelName to name of sentinel
  log "PROBE|sentinel_name|" & sentinelName
  set owned to make new presentation
  set ownedName to name of owned
  log "PROBE|owned_name|" & ownedName
  set s1 to make new slide at end of owned with properties {layout:slide layout blank}
  set s2 to make new slide at end of owned with properties {layout:slide layout blank}
  set titleBox to make new text box at end of s1 with properties {name:"parity-title", left position:40, top:40, width:560, height:60}
  set content of text range of text frame of titleBox to "Native editable PowerPoint"
  set font size of font of text range of text frame of titleBox to 28
  set bold of font of text range of text frame of titleBox to true
  set box to make new shape at end of s1 with properties {name:"parity-shape", auto shape type:autoshape rectangle, left position:40, top:140, width:200, height:90}
  log "PROBE|shape|PASS"
  set fore color of fill format of box to {32, 96, 160}
  log "PROBE|fill|PASS"
  set pic to make new picture at end of s1 with properties {name:"parity-image", file name:"/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-58ikqxox/source.png", link to file:false, save with document:true, left position:280, top:140, width:100, height:90}
  log "PROBE|picture|PASS"
  set tb to make new shape table at end of s2 with properties {name:"parity-table", number of rows:2, number of columns:2, left position:40, top:100, width:500, height:120}
  log "PROBE|table_created|PASS"
  set c to get cell from table object of tb row 1 column 1
  set content of text range of text frame of shape of c to "Native table"
  set c to get cell from table object of tb row 2 column 2
  set content of text range of text frame of shape of c to "42"
  set content of text range of text frame of shape 2 of notes page of s1 to "Native speaker note"
  set slide width of page setup of owned to 720
  log "PROBE|page_width|" & slide width of page setup of owned
  set firstId to slide ID of s1
  set secondId to slide ID of s2
  try
   set clonedSlide to duplicate s1
   if (count of slides of owned) is not 3 then error "Slide clone count mismatch"
   log clonedSlide
   if class of clonedSlide is list then set clonedSlide to item 1 of clonedSlide
   delete clonedSlide
   if (count of slides of owned) is not 2 then error "Slide remove count mismatch"
   log "PROBE|slide_clone_remove|PASS"
  on error em number en
   log "PROBE|slide_clone_remove|FAIL|" & en & "|" & em
  end try
  try
   move s2 to before s1
   if slide ID of slide 1 of owned is not secondId then error "Slide move order mismatch"
   move slide 1 of owned to after slide 2 of owned
   if slide ID of slide 1 of owned is not firstId then error "Slide move restore mismatch"
   log "PROBE|slide_move|PASS"
  on error em number en
   log "PROBE|slide_move|FAIL|" & en & "|" & em
  end try
  try
   set clonedShape to duplicate titleBox
   set name of clonedShape to "parity-clone"
   if (count of shapes of s1) is not 4 then error "Shape clone count mismatch"
   delete clonedShape
   if (count of shapes of s1) is not 3 then error "Shape remove count mismatch"
   log "PROBE|shape_clone_remove|PASS"
  on error em number en
   log "PROBE|shape_clone_remove|FAIL|" & en & "|" & em
  end try
  set left position of box to 60
  if left position of box is not 60 then error "Shape geometry edit failed"
  log "PROBE|shape_geometry_edit|PASS"
  log "PROBE|undo|BLOCKED_PRIOR_DESTRUCTIVE_UNDO"
  try
   set currentSelection to selection of document window 1 of owned
   log "PROBE|selection_type|" & (selection type of currentSelection as text)
  on error em number en
   log "PROBE|selection|FAIL|" & en & "|" & em
  end try
  log "PROBE|generation|PASS"
  save owned in pptPath as save as Open XML presentation
  set owned to presentation "native-f39221af45564258bb2b1c76221b2eed.pptx"
  log "PROBE|saved_path|" & full name of owned
  if saved of owned is false then error "Save did not mark owned presentation saved"
  if full name of owned does not end with "native-f39221af45564258bb2b1c76221b2eed.pptx" then error "Save did not bind owned output path"
  close owned saving no
  open pptAlias
  set owned to presentation "native-f39221af45564258bb2b1c76221b2eed.pptx"
  if full name of owned does not end with "native-f39221af45564258bb2b1c76221b2eed.pptx" then error "Owned reopen identity mismatch"
  log "PROBE|reopened_count|" & (count of slides of owned)
  log "PROBE|reopened_text|" & content of text range of text frame of shape "parity-title" of slide 1 of owned
  save owned in pdfPath as save as PDF
  log "PROBE|pdf|PASS"
  close owned saving no
  if saved of sentinel then error "Sentinel unexpectedly saved"
  if content of text range of text frame of shape 1 of slide 1 of sentinel is not "SENTINEL-f39221af45564258bb2b1c76221b2eed" then error "Sentinel content changed"
  repeat with prior in beforeState
   set existing to presentation (item 1 of prior)
   if {name of existing, full name of existing, saved of existing, count of slides of existing} is not contents of prior then error "Unrelated presentation changed"
  end repeat
  log "PROBE|preservation|PASS"
  return "DONE"
 end tell
end timeout
