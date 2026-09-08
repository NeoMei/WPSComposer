set f to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/tmp/wpscomposer-parity-19696ku2/native-64ea4cd440744faca8a234bb1a9518c9.pptx") as alias
tell application "Microsoft PowerPoint"
launch
open f
return count slides of presentation "native-64ea4cd440744faca8a234bb1a9518c9.pptx"
end tell