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


## Learn more

Some more useful information can be found below:
* Windows Search Samples - https://github.com/microsoft/Windows-classic-samples/tree/master/Samples/Win7Samples/winui/WindowsSearch
* Excel COM Object: https://docs.microsoft.com/en-us/office/vba/api/excel.range.findnext
* Word COM Object: https://docs.microsoft.com/en-us/office/vba/api/word.find.execute
