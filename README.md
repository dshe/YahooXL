## YahooXL&nbsp;&nbsp;[![Build status](https://ci.appveyor.com/api/projects/status/v4f5vb2g4uom43qp?svg=true)](https://ci.appveyor.com/project/dshe/yahooxl) ![GitHub release](https://img.shields.io/github/v/release/dshe/YahooXL) [![License](https://img.shields.io/badge/license-Apache%202.0-7755BB.svg)](https://opensource.org/licenses/Apache-2.0)

***Microsoft Excel add-in which retrieves data from Yahoo Finance***

- retrieves quote snapshots
- Requires [.NET Desktop Runtime 8.0](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
- Dependencies: [YahooQuotesApi](https://github.com/dshe/YahooQuotesApi), ExcelDna, NodaTime, Reactive Extensions.

### Runtime ###
  - Copy "YahooXL-AddIn64-packed.xll" (64 bit) or "YahooXL-AddIn-packed.xll" (32 bit) from GitHub Releases to a folder.
  - Double-click the file to start Excel with YahooXL loaded.
  - Open a blank spreadsheet.
  - Type in a cell: =YahooQuote()
  - Type in a cell: =YahooQuote("TSLA", "RegularMarketPrice")
  
### Installation ###
  - Download "YahooXL-AddIn64-packed.xll" (64 bit) or "YahooXL-AddIn-packed.xll" (32 bit) from GitHub Releases to a folder.
  - In Excel, with a sheet open, press `Alt+t i` to display the list of Excel Add-ins.
  - Select "Browse" to add the YahooXL add-in to the list.

### Note ###
It may be necessary to relax [Excel security settings](https://support.microsoft.com/en-us/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373). In order to avoid security warnings, [set the add-in location as trusted](https://support.microsoft.com/en-us/office/add-remove-or-change-a-trusted-location-in-microsoft-office-7ee1cdc2-483e-4cbb-bcb3-4e7c67147fb4).
