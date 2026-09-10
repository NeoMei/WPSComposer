tell application "/Applications/Microsoft Excel.app"
set targetName to "owned-958d50d0980a47d9a0f29d3b7beaa250.xlsx"
if full name of workbook targetName is not "/private/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/pytest-of-neomei/pytest-1405/test_startup_failure_persists_0/session-kl7sa9o9/owned-958d50d0980a47d9a0f29d3b7beaa250.xlsx" then error "test path mismatch"
if (count worksheets of workbook targetName) is not 1 then error "unexpected sheet count"
if value of range "A1" of worksheet 1 of workbook targetName is not "" then error "unexpected test content"
close workbook targetName saving no
if (name of every workbook) is not {"工作簿5","工作簿7"} then error "unexpected final inventory"
if saved of workbook "工作簿5" or saved of workbook "工作簿7" then error "sentinel saved state changed"
if value of range "A1" of worksheet 1 of workbook "工作簿5" is not "sentinel-58d9ceb7df8e" then error "first sentinel mismatch"
if value of range "A1" of worksheet 1 of workbook "工作簿7" is not "sentinel-e9a7d3dce1ad" then error "second sentinel mismatch"
if display alerts is not true then error "alerts changed"
return "TEST_CLEANUP_SENTINELS_PRESERVED"
end tell
