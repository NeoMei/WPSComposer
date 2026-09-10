ObjC.import('OSAKit');
ObjC.import('AppKit');
function run(argv) {
 var pid = Number(argv[0]);
 var app = $.NSRunningApplication.runningApplicationWithProcessIdentifier(pid);
 if (app.isNil() || ObjC.unwrap(app.bundleIdentifier) !== 'com.microsoft.Excel') throw Error('Excel PID mismatch');
 var source = $.NSString.stringWithContentsOfFileEncodingError(argv[1], $.NSUTF8StringEncoding, null);
 var instance = $.OSALanguageInstance.alloc.initWithLanguage($.OSALanguage.languageForName('AppleScript'));
 instance.defaultTarget = $.NSAppleEventDescriptor.descriptorWithProcessIdentifier(pid).coerceToDescriptorType(0x70736e20);
 var script = $.OSAScript.alloc.initWithSourceFromURLLanguageInstanceUsingStorageOptions(source, null, instance, 0);
 var error = Ref();
 var result = script.executeAndReturnError(error);
 if (result.isNil()) throw Error(ObjC.unwrap(error[0].description));
 return ObjC.unwrap(result.stringValue);
}
