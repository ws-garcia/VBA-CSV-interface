---
layout: default
title: Advanced Examples
parent: Examples
nav_order: 3
---

## Data subsetting

The \[EXAMPLE1\] shows how you can execute a like SQL simple query over a CSV file and dump result to a worksheet.

#### [EXAMPLE1]

```vb
Private Sub Query_CSV(path As String, ByVal keyIndex As Long, queryFilters As Variant)
    Dim CSVint As CSVinterface
    Dim CSVrecords As ECPArrayList
    Dim keyIndex As Long
    
    Set CSVint = New CSVinterface
    If path <> vbNullString Then
        Set CSVrecords = CSVint.GetCSVsubset(path, queryFilters, keyIndex) 'data filtered on keyIndex th record
        CSVint.DumpToSheet DataSource:=CSVrecords 'dump result
        Set CSVint = Nothing
        Set CSVrecords = Nothing
    End If
End Sub
```

The \[EXAMPLE2\] shows how you can split CSV data into a set of files with related data.

#### [EXAMPLE2]

```vb
Sub CSVsubSetting(path As String)
    Dim CSVint As CSVinterface
    Dim path As String
    Dim subsets As Collection

    Set CSVint = New CSVinterface
    Set subsets = CSVint.CSVsubsetSplit(path, 2) 'Subset on second field
    Set CSVint = Nothing
    Set subsets = Nothing
End Sub
```
