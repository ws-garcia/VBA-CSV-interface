---
layout: default
title: Exportation examples
parent: Examples
nav_order: 2
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

## Export data to a CSV file

The [EXAMPLE1] shows how you can export all the data in VBA array to a CSV file using the RFC-4180 standard as paramount. 

#### [EXAMPLE1]
>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>the example uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).
{: .text-grey-dk-300 .bg-grey-lt-000 }

```vb
Sub ExportToCSV_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	Dim outputFile As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	outputFile = "C:\RFC-4180_exported.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV) 'Import the range of records
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Terminate the current instance
	'@---------------------------------------------------------------------------------
	' Exportation code block start
	Call CSVix.OpenConnection(outputFile, DeleExistingFile:=True) 'Open a physical connection to the CSV file
	Call CSVix.ExportToCSV (MyArray) 'Export the array content
	Set CSVix = Nothing 'Terminate the current instance
End Sub
```

## Export to a CSV file without special syntax

The [EXAMPLE2] shows how you can export all the data in VBA array to a CSV file without check the RFC-4180 standard’s rules. Be careful, use this only if the array doesn't hold especial chars (vbCrLf [vbCr, vbLf], comma [semicolon], double quotes[apostrophe]) in neither of its fields. The output CSV file has neither field needing to be escaped.

See also
:[QuotationMode](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [EscapeType](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetype.html).

#### [EXAMPLE2]
>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>the example uses the option `QuotationMode.All`, and `EscapeType.NullChar`.
{: .text-grey-dk-300 .bg-grey-lt-000 }

```vb
Sub ExportToCSV()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	Dim outputFile As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	outputFile = "C:\RFC-4180_exported.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV) 'Import the range of records
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Terminate the current instance
	'@---------------------------------------------------------------------------------
	' Exportation code block start
	Call CSVix.OpenConnection(outputFile, DeleExistingFile:=True) 'Open a physical connection to the CSV file
	CSVix.QuotingMode = All 'Alter behavior for escaped files
	CSVix.EscapeChar = NullChar 'Specify that CSV file has neither field needing to be escaped.
	Call CSVix.ExportToCSV (MyArray) 'Export the array content
	Set CSVix = Nothing 'Terminate the current instance
End Sub
```
The [EXAMPLE3] shows how you can export all the data in VBA array to a CSV file without check the RFC-4180 standard’s rules. Each field CSV of the output file need to be escaped by desired char. The procedure presented in the [EXAMPLE3] can be used in whatever circumstance.

See also
:[QuotationMode](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [EscapeType](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetype.html).

#### [EXAMPLE3]
>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>the example uses the option `QuotationMode.All`, and `EscapeType.NullChar`.
{: .text-grey-dk-300 .bg-grey-lt-000 }

```vb
Sub ExportToCSV()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	Dim outputFile As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	outputFile = "C:\RFC-4180_exported.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV) 'Import the range of records
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Terminate the current instance
	'@---------------------------------------------------------------------------------
	' Exportation code block start
	Call CSVix.OpenConnection(outputFile, DeleExistingFile:=True) 'Open a physical connection to the CSV file
	CSVix.QuotingMode = All 'Alter behavior for escaped files
	CSVix.EscapeChar = Apostrophe 'Each CSV’s field need to be escaped with this char.
	Call CSVix.ExportToCSV (MyArray) 'Export the array content
	Set CSVix = Nothing 'Terminate the current instance
End Sub
```