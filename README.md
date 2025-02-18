ExtractMsg
---

Made this to extract .msg-files on my Mac since Outlook for Macs doesn't seem to be able to handle them. 

If you build it with

```
dotnet publish -p:PublishSingleFile=True --self-contained false --output "/<SomePath>/ExtractMsg" -r osx-arm64
```

, you'll get a single executable. Place that in some path that's on the PATH and you're good to go.

---

This first version has very little error handling and will  
write out the body-text of the .msg (and any nested .msg-s) to a text-file and export all attachments.

```
ExtractMsg will open up a .msg message.
(c) 2024 Ruben Steins - MIT License
Body text will be written as text-file and (nested) attachments exported.

Usage:
ExtractMsg filename [-d] [-f] [-h] [-v]

filename:          The name of the .msg file to extract
-f     Folder.     Create a new folder called 'filename' and extract into that.
-d     Diagnostic. Show some diagnostic messages.
-h     Help.       Shows this text.
-v     Version.    Print version info for this application.
```
---
Packages used:
[MSGReader](https://github.com/Sicos1977/MSGReader)

### Version history ###

* 1.1     [18 Feb 2025] Added support for legacy encodings, updated MsgReader to 5.7.0, added Version parameter.
* 0.1     [01 Feb 2024] [Initial release.
