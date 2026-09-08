set sourceFile to (POSIX file "/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description/docs/verification/microsoft-parity/macos-powerpoint/relocated-open/probe.pptx") as alias
with timeout of 8 seconds
 tell application "Microsoft PowerPoint"
  open sourceFile
  return count of slides of presentation "probe.pptx"
 end tell
end timeout