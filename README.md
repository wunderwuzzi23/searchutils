# searchutils
Scripts and techniques to search for through office documents, or leverage the Windows indexing service

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
