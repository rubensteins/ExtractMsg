Made this to extract .msg-files on my Mac since Outlook for Macs doesn't seem to be able to handle them. 

If you build it with

```
dotnet publish -p:PublishSingleFile=True --self-contained false --output "\<SomePath>\ExtractMsg" -r osx-arm64
```

You'll get a single executable. Place that in some path that's on the PATH and you're good to go.

This first version has no error handling whatsoever and will simply print out the body-text of the .msg (and any nested .msg-s) and export all attachments.

Packages used:
[MSGReader](https://github.com/Sicos1977/MSGReader)
