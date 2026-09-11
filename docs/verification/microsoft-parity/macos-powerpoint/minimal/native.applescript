set pptFile to POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-po3pfvrf/minimal-aea22733cf3147b2904e9d55506bb498.pptx"
set hfsPath to pptFile as text
set pdfFile to POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-po3pfvrf/native.pdf"
set pdfHFS to pdfFile as text
with timeout of 12 seconds
 tell application "Microsoft PowerPoint"
  set p to make new presentation
  set sl to make new slide at end of p with properties {layout:slide layout blank}
  set tb to make new text box at end of sl with properties {left position:40,top:40,width:500,height:70}
  set content of text range of text frame of tb to "Native basic"
  save p in hfsPath as save as Open XML presentation
  set p to presentation "minimal-aea22733cf3147b2904e9d55506bb498.pptx"
  log "SAVED|" & full name of p
  save p in pdfHFS as save as PDF
  log "PDF_SAVED"
  close p saving no
  log "CLOSED"
 end tell
 set f to pptFile as alias
 tell application "Microsoft PowerPoint"
  open f
  return "REOPENED"
 end tell
end timeout