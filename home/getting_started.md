---
layout: default
title: Getting Started
parent: Home
nav_order: 1
description: "Introduction to the VBA CSV interface class."
---

# Getting Started
{: .fs-9 }

In order to be able to use `CSVinterface.cls` within your project, please review the installation instructions.
{: .fs-6 .fw-300 }

[Install it](https://ws-garcia.github.io/VBA-CSV-interface/home/installation.html){: .btn .btn-primary .fs-5 .mb-4 .mb-md-0 .mr-2 }
---

## Usage
Import whole CSV file into an VBA array
```vbscript
Dim CSVix As CSVinterface
Dim MyArray() As String
Set CSVix = New CSVinterface
Call CSVix.OpenConnection(fileName)
Call CSVix.ImportFromCSV
Call CSVix(MyArray) 'Dumps the data to array
Set CSVix = Nothing
```
Import a range of records from a CSV file into a VBA array
```vbscript
Dim CSVix As CSVinterface
Dim MyArray As variant
Set CSVix = New CSVinterface
CSVix.StartingRecord = 10
CSVix.EndingRecord = 20
Call CSVix.OpenConnection(fileName)
Call CSVix.ImportFromCSV
Call CSVix(MyArray) 'Dumps the data to array
Set CSVix = Nothing
```
Set the char to escape special fields
```vbscript
CSVix.EscapeChar = EscapeType.NullChar
CSVix.EscapeChar = EscapeType.Apostrophe
CSVix.EscapeChar = EscapeType.DoubleQuotes
```
Set fields and records delimiters
```vbscript
CSVix.FieldsDelimiter = ";"
CSVix.RecordsDelimiter = vbCrLf
```
Define the data processing behavior
```vbscript
CSVix.QuotingMode = QuotationMode.Critical 'default
CSVix.QuotingMode = QuotationMode.All
```
Get the encoding of the last opened CSV file
```vbscript
Dim ENC as String
ENC = CSVix.FileEncoding
```
### Limitations
* __Data Format__: Keep in mind that the class doesn't distinguish between number, dates and strings, all data is read as text and you can put in an Excel sheet to let Microsoft software format it.

## Benchmark
The class was tested against two solutions (the one from [@Senipah](https://github.com/Senipah/VBA-Better-Array) and the other from [@sdkn104](https://github.com/sdkn104/VBA-CSV)) using a laptop running Win 10 Pro 64-bit, Intel® Core™ i7-4500U CPU @1.80-2.40 GHz, 8 GB RAM. 
The test consists on a fixed number of calls to the import method over three (3) different files, each of this with three records (3) and four fields (4), for an overall work load of twelve (12) fields per call:
* RFC-4180_QHO.csv: Quote Headers Only (4 fields)
* RFC-4180_HalfQ.csv: Quote Half of the fields (6 fields)
* RFC-4180_AllQ.csv: Quote All twelve (12) fields 

__NOTE: Some projects was excluded from the benchmark due they does not complies with the RFC4180 standard__.{: .fs-3 .fw-300 }

|*Procedure (Author)*|*RFC-4180_QHO.csv*|*RFC-4180_HalfQ.csv*|*RFC-4180_AllQ.csv*|
|:--------------------------|-----------------:|----------------:|----------------:|
|*ImportFromCSV (W. García)*|_N/A_|_N/A_|_N/A_|
|*FromCSV(@Senipah)*|N/A|N/A|N/A|
|*ParseCSVToArray/ADO (@sdkn104)*|N/A|N/A|N/A|

However, when setting `QuotingMode = QuotationMode.All` the class performance gets a little improve. The image below shows the performance of the VBA CSVinterface class after change the `QuotingMode` property. Take over your considerations that no all CSV files can be successful imported using the previous tweaking.

![](https://github.com/ws-garcia/VBA-CSV-interface/master/Benchmark.png)

## Licence
Copyright (C) 2020  [W. García](https://github.com/ws-garcia/VBA-CSV-interface/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.
