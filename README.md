## YahooXL&nbsp;&nbsp;[![Build status](https://ci.appveyor.com/api/projects/status/v4f5vb2g4uom43qp?svg=true)](https://ci.appveyor.com/project/dshe/yahooxl) ![GitHub release](https://img.shields.io/github/v/release/dshe/YahooXL) [![License](https://img.shields.io/badge/license-Apache%202.0-7755BB.svg)](https://opensource.org/licenses/Apache-2.0)

***Microsoft Excel add-in which retrieves data from Yahoo Finance***

- retrieves quote snapshots
- fluent interface
- Dependencies: [**ExcelDna**](https://excel-dna.net), [YahooQuotesApi](https://github.com/dshe/YahooQuotesApi), Reactive Extensions, NodaTime

### Runtime ###
  - Requires [.NET Desktop Runtime 6.0](https://dotnet.microsoft.com/en-us/download/dotnet/6.0)
  - Copy "YahooXL-AddIn64-packed.xll" (64 bit) or "YahooXL-AddIn-packed.xll" (32 bit) from GitHub Releases to a folder.
  - Double-click the file to start Excel with YahooXL loaded.
  - It may be necessary to relax [Excel security settings](https://support.microsoft.com/en-us/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373).
  - Open a blank spreadsheet.
  - Type in a cell: = YahooQuote()
  - Type in a cell: = YahooQuote("TSLA", "RegularMarketPrice")
  - Try opening the file "Test.xls", located in the solution directory.

### Installation ###
  - Copy "YahooXL-AddIn64.xll" (64 bit) or "YahooXL-AddIn.xll" (32 bit) from GitHub Releases to a folder.
  - In Excel, press Alt+t,i to display the list of Excel Add-ins.
  - Select "Browse" to add the YahooXL add-in to the list.
