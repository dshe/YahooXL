## YahooXL&nbsp;&nbsp;[![Build status](https://ci.appveyor.com/api/projects/status/v4f5vb2g4uom43qp?svg=true)](https://ci.appveyor.com/project/dshe/yahooxl) ![GitHub release](https://img.shields.io/github/v/release/dshe/YahooXL) [![License](https://img.shields.io/badge/license-Apache%202.0-7755BB.svg)](https://opensource.org/licenses/Apache-2.0)

***Microsoft Excel add-in which retrieves data from Yahoo Finance***

- retrieves quote snapshots
- fluent interface
- Requires [.NET Desktop Runtime 6](https://dotnet.microsoft.com/en-us/download/dotnet/6.0)
- Dependencies: [**ExcelDna**](https://excel-dna.net), [YahooQuotesApi](https://github.com/dshe/YahooQuotesApi), Reactive Extensions, NodaTime

### Runtime ###
  - Copy "YahooXL-AddIn64.xll" (64 bit) or "YahooXL-AddIn.xll" (32 bit) from GitHub Releases to a folder.
  - Double-click the file to start Excel with YahooXL loaded.
  - It may be necessary to relax [Excel security settings](https://support.microsoft.com/en-us/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373).
  - Open a blank spreadsheet.
  - Type in a cell: = YahooQuote()
  - Type in a cell: = YahooQuote("TSLA", "RegularMarketPrice")

### Installation ###
  - Copy "YahooXL-AddIn64.xll" (64 bit) or "YahooXL-AddIn.xll" (32 bit) from GitHub Releases to a folder.
  - In Excel, [enable the developer tab](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45).
  - On the developer tab, select "Excel Add-ins". Alternatively, press Alt+t,i to display the list of Excel Add-ins.
  - Select browse to add the YahooXL to the list.
