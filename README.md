# searchutils
Scripts and techniques to search for keywordss in Office documents, and leverage the Windows Indexing Service for quick file content searches.

## Usage

### Search-OfficeDocuments
Using the Office COM Objects this script will open Word and Excel files to search for the provided keyword. It will also look if the filename itself matches the keyword.

```
gci -r * | Search-OfficeDocuments | ft
```


### Invoke-WindowsSearch

This script directly connects to the Windows Search database and queries for the provided keyword.

```
Invoke-WindowsSearch password
```

## Long Path Names

Paths over 256 chars will produce errors, to enable long path names see the following [Microsoft article](https://learn.microsoft.com/en-us/answers/questions/1191338/windows-10-pro-22h2-enable-win32-long-path-doesnt) and/or follow these steps: 


1. Start the registry editor (regedit.exe)
2. Navigate to HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem
3. Double-click LongPathsEnabled
3. Set to 1 and click OK
4. Reboot

## Learn more

Some more useful information can be found below:
* Windows Search Samples - https://github.com/microsoft/Windows-classic-samples/tree/master/Samples/Win7Samples/winui/WindowsSearch
* Excel COM Object: https://docs.microsoft.com/en-us/office/vba/api/excel.range.findnext
* Word COM Object: https://docs.microsoft.com/en-us/office/vba/api/word.find.execute
* [Windows 10 Pro 22H2 - Enable Win32 long path doesn't work](https://learn.microsoft.com/en-us/answers/questions/1191338/windows-10-pro-22h2-enable-win32-long-path-doesnt) 
