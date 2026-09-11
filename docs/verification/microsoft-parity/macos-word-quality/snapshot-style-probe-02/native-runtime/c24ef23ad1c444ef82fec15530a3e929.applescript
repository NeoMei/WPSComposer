use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex
with timeout of 60 seconds
tell application "/Applications/Microsoft Word.app"
set nativeRows to {}
set boundDoc to document "document-85eb770d18014882a21f324998578555.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "defined-auto-update:0: set qstyle to Word style \"标题\" of boundDoc"
log probePhase
set qstyle to Word style "标题" of boundDoc
set probePhase to "defined-auto-update:1: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:2: set end of probeRows to {\"auto-update\",\"标题\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题",probeValue}
set probePhase to "defined-auto-update:3: set qstyle to Word style \"标题 1\" of boundDoc"
log probePhase
set qstyle to Word style "标题 1" of boundDoc
set probePhase to "defined-auto-update:4: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:5: set end of probeRows to {\"auto-update\",\"标题 1\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 1",probeValue}
set probePhase to "defined-auto-update:6: set qstyle to Word style \"标题 2\" of boundDoc"
log probePhase
set qstyle to Word style "标题 2" of boundDoc
set probePhase to "defined-auto-update:7: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:8: set end of probeRows to {\"auto-update\",\"标题 2\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 2",probeValue}
set probePhase to "defined-auto-update:9: set qstyle to Word style \"标题 3\" of boundDoc"
log probePhase
set qstyle to Word style "标题 3" of boundDoc
set probePhase to "defined-auto-update:10: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:11: set end of probeRows to {\"auto-update\",\"标题 3\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 3",probeValue}
set probePhase to "defined-auto-update:12: set qstyle to Word style \"标题 4\" of boundDoc"
log probePhase
set qstyle to Word style "标题 4" of boundDoc
set probePhase to "defined-auto-update:13: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:14: set end of probeRows to {\"auto-update\",\"标题 4\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 4",probeValue}
set probePhase to "defined-auto-update:15: set qstyle to Word style \"标题 5\" of boundDoc"
log probePhase
set qstyle to Word style "标题 5" of boundDoc
set probePhase to "defined-auto-update:16: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:17: set end of probeRows to {\"auto-update\",\"标题 5\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 5",probeValue}
set probePhase to "defined-auto-update:18: set qstyle to Word style \"标题 6\" of boundDoc"
log probePhase
set qstyle to Word style "标题 6" of boundDoc
set probePhase to "defined-auto-update:19: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:20: set end of probeRows to {\"auto-update\",\"标题 6\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 6",probeValue}
set probePhase to "defined-auto-update:21: set qstyle to Word style \"标题 7\" of boundDoc"
log probePhase
set qstyle to Word style "标题 7" of boundDoc
set probePhase to "defined-auto-update:22: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:23: set end of probeRows to {\"auto-update\",\"标题 7\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 7",probeValue}
set probePhase to "defined-auto-update:24: set qstyle to Word style \"标题 8\" of boundDoc"
log probePhase
set qstyle to Word style "标题 8" of boundDoc
set probePhase to "defined-auto-update:25: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:26: set end of probeRows to {\"auto-update\",\"标题 8\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 8",probeValue}
set probePhase to "defined-auto-update:27: set qstyle to Word style \"标题 9\" of boundDoc"
log probePhase
set qstyle to Word style "标题 9" of boundDoc
set probePhase to "defined-auto-update:28: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:29: set end of probeRows to {\"auto-update\",\"标题 9\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","标题 9",probeValue}
set probePhase to "defined-auto-update:30: set qstyle to Word style \"副标题\" of boundDoc"
log probePhase
set qstyle to Word style "副标题" of boundDoc
set probePhase to "defined-auto-update:31: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:32: set end of probeRows to {\"auto-update\",\"副标题\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","副标题",probeValue}
set probePhase to "defined-auto-update:33: set qstyle to Word style \"列表段落\" of boundDoc"
log probePhase
set qstyle to Word style "列表段落" of boundDoc
set probePhase to "defined-auto-update:34: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:35: set end of probeRows to {\"auto-update\",\"列表段落\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","列表段落",probeValue}
set probePhase to "defined-auto-update:36: set qstyle to Word style \"明显参考\" of boundDoc"
log probePhase
set qstyle to Word style "明显参考" of boundDoc
set probePhase to "defined-auto-update:37: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:38: set end of probeRows to {\"auto-update\",\"明显参考\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","明显参考",probeValue}
set probePhase to "defined-auto-update:39: set qstyle to Word style \"明显强调\" of boundDoc"
log probePhase
set qstyle to Word style "明显强调" of boundDoc
set probePhase to "defined-auto-update:40: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:41: set end of probeRows to {\"auto-update\",\"明显强调\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","明显强调",probeValue}
set probePhase to "defined-auto-update:42: set qstyle to Word style \"明显引用\" of boundDoc"
log probePhase
set qstyle to Word style "明显引用" of boundDoc
set probePhase to "defined-auto-update:43: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:44: set end of probeRows to {\"auto-update\",\"明显引用\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","明显引用",probeValue}
set probePhase to "defined-auto-update:45: set qstyle to Word style \"默认段落字体\" of boundDoc"
log probePhase
set qstyle to Word style "默认段落字体" of boundDoc
set probePhase to "defined-auto-update:46: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:47: set end of probeRows to {\"auto-update\",\"默认段落字体\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","默认段落字体",probeValue}
set probePhase to "defined-auto-update:48: set qstyle to Word style \"普通表格\" of boundDoc"
log probePhase
set qstyle to Word style "普通表格" of boundDoc
set probePhase to "defined-auto-update:49: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:50: set end of probeRows to {\"auto-update\",\"普通表格\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","普通表格",probeValue}
set probePhase to "defined-auto-update:51: set qstyle to Word style \"无列表\" of boundDoc"
log probePhase
set qstyle to Word style "无列表" of boundDoc
set probePhase to "defined-auto-update:52: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:53: set end of probeRows to {\"auto-update\",\"无列表\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","无列表",probeValue}
set probePhase to "defined-auto-update:54: set qstyle to Word style \"引用\" of boundDoc"
log probePhase
set qstyle to Word style "引用" of boundDoc
set probePhase to "defined-auto-update:55: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:56: set end of probeRows to {\"auto-update\",\"引用\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","引用",probeValue}
set probePhase to "defined-auto-update:57: set qstyle to Word style \"正文\" of boundDoc"
log probePhase
set qstyle to Word style "正文" of boundDoc
set probePhase to "defined-auto-update:58: set probeValue to automatically update of qstyle"
log probePhase
set probeValue to automatically update of qstyle
set probePhase to "defined-auto-update:59: set end of probeRows to {\"auto-update\",\"正文\",probeValue}"
log probePhase
set end of probeRows to {"auto-update","正文",probeValue}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
