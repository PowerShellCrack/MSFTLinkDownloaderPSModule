# LGPO PowerShell module

A module that will download files from Microsoft Link ID's

## Prerequisites

None

## Cmdlets
- **Get-MicrosoftProduct** - Downloads file from linkID

## Install

```powershell
Install-Module MSFTLinkDownloader -Force
Import-Module MSFTLinkDownloader
```

## Examples

```powershell
.EXAMPLE
Get-MicrosoftProduct -LinkID '49117' -DestPath C:\temp\Downloads -Force

.EXAMPLE
Get-MicrosoftProduct -LinkID '55319' -DestPath C:\temp\Downloads -Filter 'LGPO'

.EXAMPLE
Get-MicrosoftProduct -LinkID '49117' -DestPath C:\temp\Downloads -Force -Extract -Cleanup

.EXAMPLE
Get-MicrosoftProduct -LinkID '55319' -DestPath C:\temp\Downloads -Passthru

.EXAMPLE
49117,55319,104223 | Get-MicrosoftProduct -DestPath C:\temp\Downloads -Passthru

.EXAMPLE
49117,104223 | Get-MicrosoftProduct -DestPath C:\temp\Downloads -Passthru -NoProgress -Extract -Cleanup

.EXAMPLE
'55319' | Get-MicrosoftProduct -DestPath C:\temp\Downloads -Filter 'Windows Server' -Passthru -Extract -Verbose

```

## Passthru

The output when using _-Passthru_ parameter will create object data that can be further used in a pipeline or script

![Output](/.images/PassthruData.jpg)
