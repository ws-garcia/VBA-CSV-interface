---
layout: default
title: Importation examples
parent: Examples
nav_order: 1
---

## Import CSV file data

The \[EXAMPLE1\] shows how you can import all the data from a CSV file. 

#### [EXAMPLE1]

```vb
Sub ImportRecords()
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
	Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE2\] shows how you can import all the data from a CSV file using Dynamic Typing. 

#### [EXAMPLE2]

```vb
Sub TEST_DynamicTyping()
	Dim conf As parserConfig
	Dim CSVstring As String
	Dim Arr() As Variant
	
	Set CSVint = New CSVinterface
	Set conf = New parserConfig
	With conf
		.recordsDelimiter = vbCrLf
		.path = "C:\100000.quoted.csv"
		.dynamicTyping = True
		.defineTypingTemplate TypeConversion.ToDate, _
                            TypeConversion.ToLong, _
                            TypeConversion.ToDate, _
                            TypeConversion.ToLong, _
                            TypeConversion.ToDouble, _
                            TypeConversion.ToDouble, _
                            TypeConversion.ToDouble
		.defineTypingTemplateLinks 6, _
                                 7, _
                                 8, _
                                 9, _
                                 10, _
                                 11, _
                                 12
	End With
	CSVint.GuessDelimiters conf 'Try to guess CSV file data delimiters
	CSVint.ImportFromCSV(conf).DumpToArray Arr 'Import and dump the data to an array
	Set CSVint = Nothing
End Sub
```

With the \[EXAMPLE3\] you will learn how to mimic a SQL statement and dump the resulting data object to an Excel spreadsheet.

The SQL query is:
SELECT *
FROM Demo_100k
WHERE (((Demo_100k.\[Region\])='Central America and the Caribbean') AND ((Demo_100k.\[Sales_Channel\])='Online'))
ORDER BY Demo_100k.\[Order_Date\] DESC;

#### [EXAMPLE3]
```vb
Public Sub SQL_Param_Query()
'This example shows how to filter CSV content by "Region" and "Sales_Channel", sort data and dump to an Excel spreadsheet.
    Dim path As String
    Dim conf As parserConfig
    Dim CSVrecord As ECPArrayList
    Dim CSVrecords As ECPArrayList
    
    Set CSVint = New CSVinterface
    Set conf = CSVint.parseConfig
    Set CSVrecords = New ECPArrayList
    With conf
        .recordsDelimiter = vbCr
        path = "C:\"C:\100000.quoted.csv"
        .path = path
        .dynamicTyping = True
        .headers = True
        .DefineTypingTemplate TypeConversion.ToDate, _
                                TypeConversion.ToLong, _
                                TypeConversion.ToDate, _
                                TypeConversion.ToLong, _
                                TypeConversion.ToDouble, _
                                TypeConversion.ToDouble, _
                                TypeConversion.ToDouble
        .DefineTypingTemplateLinks 6, _
                                7, _
                                8, _
                                9, _
                                10, _
                                11, _
                                12
    End With
	 '@---------------------------------------------------------------------------
	 ' Sequential reading
    CSVint.OpenSeqReader conf 'Open a Sequential Reader
    Set CSVrecord = CSVint.GetRecord 'Get CSV record
    If conf.headers Then
        If Not CSVrecord Is Nothing Then
            CSVrecords.Add CSVrecord(0) 'Save the CSV header
        End If
    End If
    Do While Not CSVrecord Is Nothing 'Loop to EOF or error occurrence
        If CSVrecord(0)(0) = "Central America and the Caribbean" And CSVrecord(0)(3) = "Online" Then
            CSVrecords.Add CSVrecord(0) 'Append an One-dimentional array
        End If
        Set CSVrecord = CSVint.GetRecord 'Load next CSV record
    Loop
    If Not CSVint.importSuccess Then
        If MsgBox("The operation was aborted due to an error. Ignore error and dump the data?", vbYesNo, "Alert") = vbYes Then
            CSVrecords.Sort 2, SortColumn:=6, Descending:=True
            CSVint.DumpToSheet DataSource:=CSVrecords
            '@---------------------------------------------------------------------------
            ' Alternative way:
            '                   Set CSVint.items = CSVrecords
            '                   CSVint.Sort(SortColumn:=6, Descending:=True).DumpToSheet
            '@---------------------------------------------------------------------------
        
        End If
    Else
        CSVrecords.Sort 2, SortColumn:=6, Descending:=True
        CSVint.DumpToSheet DataSource:=CSVrecords
    End If
    Set CSVint = Nothing
    Set CSVrecords = Nothing
End Sub
```

The \[EXAMPLE4\] shows how you can dump the imported data to an Excel Worksheet.
#### [EXAMPLE4]
```vb
Sub ImportAndDumpToSheet()
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
	CSVint.ImportFromCSV(conf).DumpToSheet 'Import and dump the data to a new Worksheet
	Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE5\] shows how you can dump the imported data to an Access Database. The created table will have some indexed fields.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>This method is only available in the [Access version of the CSVinterface.cls](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip) module.
{: .text-grey-dk-300 .bg-yellow-000 }

#### [EXAMPLE5]
```vb
Sub ImportAndDumpToAccessDB()
	Dim path As String
	Dim conf As parserConfig
	Dim dBase As DAO.Database
	
	Set CSVint = New CSVinterface
	Set conf = CSVint.ParseConfig
	With conf
	    .path = "C:\100000.quoted.csv"
	    .dynamicTyping = False
	End With
	CSVint.GuessDelimiters conf 'Try to guess CSV file data delimiters
	Set dBase = CurrentDb
	'Import and dump the data into a new database table. This will create indexes for the "Region" field and for the second field in the table.
	CSVint.ImportFromCSV(conf).DumpToAccessTable dBase, "CSV_ImportedData", "Region", 2
	Set CSVint = Nothing
	Set dBase = Nothing
End Sub
```