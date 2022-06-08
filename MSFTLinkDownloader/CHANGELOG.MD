# Change log for MSFTLinkDownloader module

## 1.0.0.3 June 08, 2022

- Changed Get-MSFTLink to just link retrieval; replaced with Invoke-MSFTLinkDownload
- Added pipeline from  Get-MSFTLink to Invoke-MSFTLinkDownload; simplifies download

## 1.0.0.2 June 06, 2022

- Fixed parameter positions on cmdlets; set first one to 0
- Set progress bare to completed when done; hide progress when done.


## 1.0.0.1 June 05, 2022

- Changed cmdlet name; unable to use "Microsoft" in cmdlet because it's not official
- Removed any reference to Microsoft in man

## 1.0.0.0 June 04, 2022

- Initial upload
