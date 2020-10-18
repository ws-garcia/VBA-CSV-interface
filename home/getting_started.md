---
layout: default
title: Getting Started
parent: Home
nav_order: 1
description: "Introduction to the VBA CSV interface class."
---

# Getting Started

In order to be able to use `CSVinterface.cls` within your project, please review the [installation instructions](https://ws-garcia.github.io/VBA-CSV-interface/home/installation.html).
---

{: .no_toc }

<details open markdown="block">
  <summary>
    Table of contents
  </summary>
  {: .text-delta }
1. TOC
{:toc}
</details>

The CSV files are special kind of tabulated plain text data widely used in data exchange. There is no globally accepted standard format for that kind of files, however, out there are well formed standards such as [RFC4180](https://www.ietf.org/rfc/rfc4180.txt) proposed by The Internet Society.
Although many solutions has been developed for work with CSV files into VBA, including projects from [@sdkn104](https://github.com/sdkn104/VBA-CSV) and [@Senipah](https://github.com/Senipah/VBA-Better-Array) on GitHub, the performance philosophy conduce me to the development of a VBA class module that allows users exchange data between VBA arrays and CSV files at superior speed for the VBA programing language.

## Advantages
* Fully compliant with RFC4180 CSV standard.
* Exported data is 100% Excel spreadsheet compatible.
* Writes and reads files at high speed.
* Minimal Memory overload.
* User have the option to import only certain range of records from given CSV file.
* Auto exclude any quote mark when data is imported.
* Simple code logic that allows you easy modify and enhance it!

## Philosophy
The VBA CSVinterface class module is designed for gain advantage from the well-structured CSV files, this means, there isn't automatic syntax check, given the user decide how the class will works. This can be seen as a weakness, but the class get a speed-up on writing and reading procedures at time the user controls how the file is interpreted, keeping in mind that, in fact, VBA is a language with slow code execution speed. Under this idealization the developed solution complies with the RFC4180 standard for user specified CSV document format.

Keep in mind that class intentionally ignores the rule #7 of the RFC4180 standard [If double-quotes are used to enclose fields, then a double-quote appearing inside a field must be escaped by preceding it with another double quote], letting the users the chance to define the apostrophe as escape char.
{: .fs-4 .fw-300 }

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
