## YahooXL&nbsp;&nbsp;[![Build status](https://ci.appveyor.com/api/projects/status/v4f5vb2g4uom43qp?svg=true)](https://ci.appveyor.com/project/dshe/yahooxl) ![GitHub release](https://img.shields.io/github/v/release/dshe/YahooXL) [![License](https://img.shields.io/badge/license-Apache%202.0-7755BB.svg)](https://opensource.org/licenses/Apache-2.0)

***Microsoft Excel add-in which retrieves data from Yahoo Finance.***

- retrieves quote snapshots
- fluent interface
- tested

Dependencies: YahooQuotesApi, ExcelDna, Reactive Extensions, NodaTime.

Runtime: 
  - Download Package.zip from the GitHub release. 
  - Extract files to a folder.
  - Execute "YahooXL-AddIn64.xll" to start Microsoft Excel.
  - Open a blank spreadsheet.
  - Type in a cell: =YahooQuote("C", "RegularMarketPrice")
  
  It may be required to relax Excel macro and protected view security settings.
