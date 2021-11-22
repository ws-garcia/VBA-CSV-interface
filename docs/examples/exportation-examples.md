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
    Dim conf As CSVparserConfig

    Set CSVint = New CSVinterface
    Set conf = CSVint.parseConfig
    With conf
        .path = "C:\100000.quoted.csv"
        .dynamicTyping = False
        Set .dialect = CSVint.SniffDelimiters(conf)   'Try to guess CSV file data delimiters\
        CSVint.ImportFromCSV conf 'Import the data
        .path = Environ("USERPROFILE") & "\Desktop\100000.quoted.tsv"
        .dialect.fieldsDelimiter = vbTab
        CSVint.ExportToCSV CSVint.items, conf 'Export internal items
    End With
    Set CSVint = Nothing 'Terminate the current instance
End Sub
```

The \[EXAMPLE2\] shows how you can import a CSV file, sort the data, dump to a VBA array and export to a new CSV file.

#### [EXAMPLE2]

```vb
Sub SortAndExportToCSV()
    Dim CSVint As CSVinterface
    Dim conf As CSVparserConfig

    Set CSVint = New CSVinterface
    Set conf = CSVint.parseConfig
    With conf
        .path = "C:\100000.quoted.csv"
        .Headers = True 'The header will not sorted
        .dynamicTyping = False
        Set .dialect = CSVint.SniffDelimiters(conf) 'Try to guess CSV file data delimiters
        CSVint.ImportFromCSV(conf).Sort SortingKeys:=1 'Import and sort the data in ascending way
        .path = Environ("USERPROFILE") & "\Desktop\100000.quoted.tsv"
        CSVint.ExportToCSV CSVint.items, conf 'Export internal items
    End With
    Set CSVint = Nothing 'Terminate the current instance
End Sub
```