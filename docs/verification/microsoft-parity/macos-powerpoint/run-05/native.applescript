with timeout of 30 seconds
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
  set content of text range of text frame of sentText to "SENTINEL-fc4ba1a2aea1497b8839179cdb7c73fb"
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
  -- Picture insertion tested separately; prior native runs retained its timeout.
  set imagePath to "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/run-05/source.png"
  log "PROBE|picture|BLOCKED_PRIOR_TIMEOUT"
  set tb to make new shape table at end of s2 with properties {name:"parity-table", number of rows:2, number of columns:2, left position:40, top:100, width:500, height:120}
  log "PROBE|table_created|PASS"
  set c to get cell from table object of tb row 1 column 1
  set content of text range of text frame of shape of c to "Native table"
  set c to get cell from table object of tb row 2 column 2
  set content of text range of text frame of shape of c to "42"
  set content of text range of text frame of shape 2 of notes page of s1 to "Native speaker note"
  log "PROBE|generation|PASS"
  save owned in "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/run-05/native.pptx" as save as Open XML presentation
  log "PROBE|saved_path|" & full name of owned
  close owned saving no
  open "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/run-05/native.pptx"
  set owned to presentation "native.pptx"
  if full name of owned does not end with "native.pptx" then error "Owned reopen identity mismatch"
  log "PROBE|reopened_count|" & (count of slides of owned)
  log "PROBE|reopened_text|" & content of text range of text frame of shape "parity-title" of slide 1 of owned
  save owned in "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/run-05/native.pdf" as save as PDF
  log "PROBE|pdf|PASS"
  close owned saving no
  if saved of sentinel then error "Sentinel unexpectedly saved"
  if content of text range of text frame of shape 1 of slide 1 of sentinel is not "SENTINEL-fc4ba1a2aea1497b8839179cdb7c73fb" then error "Sentinel content changed"
  repeat with prior in beforeState
   set existing to presentation (item 1 of prior)
   if {name of existing, full name of existing, saved of existing, count of slides of existing} is not contents of prior then error "Unrelated presentation changed"
  end repeat
  log "PROBE|preservation|PASS"
  return "DONE"
 end tell
end timeout
