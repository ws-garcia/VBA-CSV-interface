---
layout: default
title: Exportation examples
parent: Examples
nav_order: 2
---

## Export data to a CSV file

The \[EXAMPLE1\] shows how you can import a CSV file, dump the data to a VBA array and export to a new TSV file.

#### [EXAMPLE1]

```vb
Sub ExportToTSV()
	Dim CSVint As CSVinterface
	Dim conf As parserConfig
	Dim Arr() As Variant

	Set CSVint = New CSVinterface
	Set conf = CSVint.ParseConfig
	With conf
	    .path = "C:\100000.quoted.csv"
	    .dynamicTyping = False
	End With
	CSVint.GuessDelimiters conf 'Try to guess CSV file data delimiters
	CSVint.ImportFromCSV(conf).DumpToArray Arr 'Import and dump the data to an array
	With conf
	    .path = Environ("USERPROFILE") & "\Desktop\100000.quoted.tsv"
	    .fieldsDelimiter = vbTab
	End With
	CSVint.ExportToCSV Arr, conf
	Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE2\] shows how you can import a CSV file, sort the data, dump to a VBA array and export to a new CSV file.

#### [EXAMPLE2]

```vb
Sub SortAndExportToCSV()
	Dim CSVint As CSVinterface
	Dim conf As parserConfig
	Dim Arr() As Variant

	Set CSVint = New CSVinterface
	Set conf = CSVint.ParseConfig
	With conf
	    .path = "C:\100000.quoted.csv"
		 .headers = True 'The header will not sorted
	    .dynamicTyping = False
	End With
	CSVint.GuessDelimiters conf 'Try to guess CSV file data delimiters
	CSVint.ImportFromCSV(conf).Sort(SortColumn:=1, Descending:=False).DumpToArray Arr 'Import, sort and dump the data to an array
	conf.path = Environ("USERPROFILE") & "\Desktop\100000.quoted.tsv"
	CSVint.ExportToCSV Arr, conf
	Set CSVint = Nothing 'Terminate the current instance
End Sub
```