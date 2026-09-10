import Foundation
import ScriptingBridge
import AppKit
let pid = Int32(CommandLine.arguments[1])!
guard let running = NSRunningApplication(processIdentifier: pid), running.bundleIdentifier == "com.microsoft.Excel", let app = SBApplication(processIdentifier: pid) else { exit(2) }
app.timeout = 120
let name = app.value(forKey: "name") ?? NSNull()
let alerts = app.value(forKey: "displayAlerts") ?? NSNull()
guard let books = app.value(forKey: "workbooks") as? SBElementArray else { exit(3) }
var inventory: [[String:Any]] = []
for book in books {
 guard let book = book as? SBObject else { exit(4) }
 inventory.append(["name": book.value(forKey: "name") ?? NSNull(), "path": book.value(forKey: "fullName") ?? NSNull(), "saved": book.value(forKey: "saved") ?? NSNull()])
}
let result: [String:Any] = ["pid":pid,"bundle":running.bundleIdentifier!,"name":name,"alerts":alerts,"running":app.isRunning,"workbooks":inventory]
let data = try JSONSerialization.data(withJSONObject: result, options: [.prettyPrinted, .sortedKeys])
print(String(data:data,encoding:.utf8)!)
