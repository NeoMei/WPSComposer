tell application "Microsoft PowerPoint"
set p to presentation "probe.pptx"
if full name of p is not "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/relocated-open/probe.pptx" then error "wrong document"
if saved of p is false then error "document has unsaved changes"
close p saving no
return "OWNED_CLOSED"
end tell