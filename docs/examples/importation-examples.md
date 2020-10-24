---
layout: default
title: Importation examples
parent: Examples
nav_order: 1
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

#### Import CSV file data

The *example1* shows how you can import all the data from a CSV file using the RFC-4180 standard as paramount. 

###### [example1]
*Note: the example uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	Call CSVix.OpenConnection(fileName) 'Open a physical connection to the CSV file
	Call CSVix.ImportFromCSV 'Import data
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```

#### Import top 10 records from a CSV file into a VBA array 

The *example2* shows how you can import the Top 10 records from a CSV file using the RFC-4180 standard as paramount.

###### [example2]
*Note: the example uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.OpenConnection(fileName) 'Open a physical connection to the CSV file
	Call CSVix.ImportFromCSV 'Import the range of records
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```
The *example3* accomplishes the same task of the *example1*, the difference is that a temporary variable is used to store the CSV file's content instead of use the `OpenConnection` method. Also, the *example3* shows how to omit the CSV's headers.

###### [example3]
*Note: the example uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.EndingRecord = 10 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV, HeadersOmission:=True) 'Import the range of records omitting the headers
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```

#### Import ten middle records from a CSV file into a VBA array 
The *example4* shows how you can import 10 middle records from a CSV file using the RFC-4180 standard as paramount.

###### [example4]
*Note: the example uses the option `QuotationMode.Critical`, [learn more here](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String, tmpCSV As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	tmpCSV = CSVix.GetDataFromCSV(filePath) 'Store CSV file's content.
	CSVix.StartingRecord = 11 'Sets the importation ending
	CSVix.EndingRecord = 20 'Sets the importation ending
	Call CSVix.ImportFromCSVString(tmpCSV, HeadersOmission:=True) 'Import the range of records omitting the headers
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```

#### Import CSV file that haven’t special syntax

This is the fastest way to work with CSV files because the CSV interface class don't check the syntax given at the RFC-4180 standard. If your CSV files has trailing spaces, or you don't know if it holds a field needing to be escaped, please [reset the config options](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/resettodefault.html) for the current instance to avoid incorrect results.

The *example5* shows how you can import all the data from a CSV file without checking the syntax given at the RFC-4180 standard. The file to be parsed has neither field needing to be escaped.

See also
:[QuotationMode](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [EscapeType](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetype.html).
###### [example5]
*Note: the example uses the option `QuotationMode.All`, and `EscapeType.NullChar`*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String
	
	filePath = "C:\Demo_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	Call CSVix.OpenConnection(fileName) 'Open a physical connection to the CSV file
	CSVix.QuotingMode = All 'Alter behavior for escaped files
	CSVix.EscapeChar = NullChar 'Specify that CSV file has neither field needing to be escaped.
	Call CSVix.ImportFromCSV 'Import data
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```
The *example6* shows how you can import all the data from a CSV file without checking the syntax given at the RFC-4180 standard. In the file to be parsed, all fields need to be escaped.

See also
:[QuotationMode](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [EscapeType](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetype.html).
###### [example6]
*Note: the example uses the option `QuotationMode.All`, and `EscapeType.DoubleQuotes`*
```vb
Sub ImportTopTenRecords_RFC4180()
	Dim CSVix As CSVinterface
	Dim MyArray() As String
	Dim filePath As String
	
	filePath = "C:\Demo_Headed_400k_records.csv" 'Change this to suit your needs
	Set CSVix = New CSVinterface 'Create new instance
	Call CSVix.OpenConnection(fileName) 'Open a physical connection to the CSV file
	CSVix.QuotingMode = All 'Alter behavior for escaped files
	CSVix.EscapeChar = DoubleQuotes 'Specify that all fields need to be escaped.
	Call CSVix.ImportFromCSV 'Import data
	Call CSVix(MyArray) 'Dumps the data to array
	Set CSVix = Nothing 'Close the current instance
End Sub
```