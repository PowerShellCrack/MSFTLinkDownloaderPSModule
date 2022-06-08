# LGPO PowerShell module

A module that will download files from Microsoft Link ID's

## Prerequisites

None

## Cmdlets
- **Get-MSFTLink** - Get the download url from linkID; export to object
- **Invoke-MSFTLinkDownload** - Downloads file from linkID

## Install

```powershell
Install-Module MSFTLinkDownloader -Force
Import-Module MSFTLinkDownloader
```

## Examples

```powershell

#grab linkID download URLS
Get-MSFTLink -LinkID '49117'

#grab linkID download URLS with LGPO in name
Get-MSFTLink -LinkID '55319' -Filter 'LGPO'

#grab linkID download URLS for British english language
49117,55319,104223 | Get-MSFTLink -Language en-gb

#download file by Link ID with Server in name (overwrite if exists)
Invoke-MsftLinkDownload -LinkID 55319,104223 -Filter 'Server' -DestPath C:\temp\Downloads -Force

#download file by URL (overwrite if exists)
Invoke-MsftLinkDownload -DownloadLink 'https://download.microsoft.com/download/8/5/C/85C25433-A1B0-4FFA-9429-7E023E7DA8D8/LGPO.zip' -DestPath C:\temp\Downloads -Force

#grab linkID download URLS, then download them and export their status
Get-MSFTLink -LinkID 49117,104223 | Invoke-MsftLinkDownload -DestPath C:\temp\Downloads -Passthru

#collect linkID data, then download them, show no progress bar and extract them if they are archive files as well a delete the archive when done.
$Links = Get-MSFTLink -LinkID 49117,55319,104223
$Links | Invoke-MsftLinkDownload -DestPath C:\temp\Downloads -Passthru -NoProgress -Extract -Cleanup

```

## Passthru

The output when using _-Passthru_ parameter will create object data that can be further used in a pipeline or script

![Output](/.images/PassthruData.jpg)
