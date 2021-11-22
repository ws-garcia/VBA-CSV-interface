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
    Dim Arr() As Variant

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\100000.quoted.csv"
    End With
    With CSVint
        .SniffDelimiters .parseConfig 'Try to guess CSV file data delimiters
        .ImportFromCSV(.parseConfig).DumpToArray Arr 'Import and dump the data to an array
    End With
    Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE2\] shows how you can import all the data from a CSV file using Dynamic Typing. 

#### [EXAMPLE2]

```vb
Sub TEST_DynamicTyping()
    Dim CSVint As CSVinterface
    Dim CSVstring As String
    Dim Arr() As Variant
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .dialect.recordsDelimiter = vbCrLf
        .path = "C:\100000.quoted.csv"
        .dynamicTyping = True
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
    With CSVint
        .SniffDelimiters .parseConfig 'Try to guess CSV file data delimiters
        .ImportFromCSV(.parseConfig).DumpToArray Arr 'Import and dump the data to an array
    End With
    Set CSVint = Nothing
End Sub
```

The \[EXAMPLE3\] shows how you can dump the imported data to an Excel Worksheet.

#### [EXAMPLE3]
```vb
Sub ImportAndDumpToSheet()
    Dim CSVint As CSVinterface
    Dim conf As CSVparserConfig

    Set CSVint = New CSVinterface
    Set conf = CSVint.parseConfig
    With conf
        .path = "C:\100000.quoted.csv"
    End With
    With CSVint
        .SniffDelimiters conf 'Try to guess CSV file data delimiters
        .ImportFromCSV(conf).DumpToSheet 'Import and dump the data to a new Worksheet
    End With
    Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE4\] shows how you can dump the imported data to an Access Database. The created table will have some indexed fields.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>This method is only available in the [Access version of the CSVinterface.cls](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip) module.
{: .text-grey-dk-300 .bg-yellow-000 }

#### [EXAMPLE4]
```vb
Sub ImportAndDumpToAccessDB()
    Dim CSVint As CSVinterface
    Dim path As String
    Dim dBase As DAO.Database
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\100000.quoted.csv"
    End With
    With CSVint
        .SniffDelimiters .parseConfig 'Try to guess CSV file data delimiters
        Set dBase = CurrentDb
        'Import and dump the data into a new database table. This will create indexes for the "Region" field and for the second field in the table.
        .ImportFromCSV(.parseConfig).DumpToAccessTable dBase, "CSV_ImportedData", "Region", 2
    End With
    Set CSVint = Nothing
    Set dBase = Nothing
End Sub
```

The \[EXAMPLE5\] shows how you can loop, **one by one**, through all available records in a CSV file using the sequential reader.

#### [EXAMPLE5]
```vb
Sub SequentialImport()
    Dim CSVint As CSVinterface
    Dim csvRecord As CSVArrayList
            
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\100000.quoted.csv"
    End With
    With CSVint
        .SniffDelimiters .parseConfig
        .OpenSeqReader .parseConfig
        Do
            Set csvRecord = .GetRecord
        Loop While Not csvRecord Is Nothing
    End With
    Set CSVint = Nothing
End Sub
```

The \[EXAMPLE6\] shows how you can loop, **set by set**, through all available records in a CSV file using the `ECPTextStream` class module.

#### [EXAMPLE6]
```vb
Sub ImportCSVinChunks()
    Dim CSVint As CSVinterface
    Dim StreamReader As CSVTextStream
            
    Set CSVint = New CSVinterface
    With CSVint
        .parseConfig.path = "C:\Sample.csv"
        .SniffDelimiters .parseConfig                       ' Try to guess delimiters
    End With
    Set StreamReader = New CSVTextStream
    With StreamReader
        .endStreamOnLineBreak = True                        ' Instruct to find line breaks
        .OpenStream CSVint.parseConfig.path                 ' Connect to CSV file
        Do
            .ReadText                                       ' Read a CSV chunk
            CSVint.ImportFromCSVString .bufferString, _
                                    CSVint.parseConfig      ' Import a set of records
        Loop While Not .atEndOfStream                       ' Continue until reach the end of the CSV file.
    End With
    Set CSVint = Nothing
    Set StreamReader = Nothing
End Sub
```