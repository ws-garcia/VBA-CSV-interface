---
layout: default
title: File Conversion
parent: Examples
nav_order: 3
---

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>All the examples uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).
{: .text-grey-dk-300 .bg-grey-lt-000 }

## Convert a CSV file to TSV

The \[EXAMPLE1\] shows how you can turn a CSV file to TSV. 

#### [EXAMPLE1]

```vb
Sub ExportToCSV_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	Dim outputFile As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	outputFile = "C:\Demo_400k_records.tsv" 'Change this to suit your needs
	
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV) 'Import the range of records
	Call CSVix.DumpToArray(MyArray) 'Dumps the data to array
	'@---------------------------------------------------------------------------------
	' Exportation code block start
	CSVix.FieldsDelimiter = vbTab
	Call CSVix.OpenConnection(outputFile, DeleExistingFile:=True) 'Open a physical connection to the TSV file
	Call CSVix.ExportToCSV (MyArray) 'Export the array content
	Set CSVix = Nothing 'Terminate the current instance
End Sub
```

## Convert a TSV file to CSV

The \[EXAMPLE2\] shows how you can turn a TSV file to CSV. 

#### [EXAMPLE2]

```vb
Sub ExportToCSV_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	Dim outputFile As String
	
	filePath = "C:\Demo_400k_records.tsv" 'Change this to suit your needs
	outputFile = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store TSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV) 'Import the range of records
	Call CSVix.DumpToArray(MyArray) 'Dumps the data to array
	'@---------------------------------------------------------------------------------
	' Exportation code block start
	CSVix.FieldsDelimiter = ","
	Call CSVix.OpenConnection(outputFile, DeleExistingFile:=True) 'Open a physical connection to the CSV file
	Call CSVix.ExportToCSV (MyArray) 'Export the array content
	Set CSVix = Nothing 'Terminate the current instance
End Sub
```