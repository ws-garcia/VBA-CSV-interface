---
title: Dedupe
parent: Methods
grand_parent: API
nav_order: 4
---

# Dedupe
{: .d-inline-block }

New
{: .label .label-purple }

Returns a list of records as a result of the deduplication of the imported CSV data.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`Dedupe`*(keys)*

### Parameters

The required *keys* argument is an identifier specifying a `String` Type variable. Represents the indexes of the fields used for deduplication.

### Returns value

*Type*: `CSVArrayList`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html).

## Behavior

The `keys` parameter will indicate which fields/columns will be used in the deduplication. A string like `"0,5"` used as parameter will deduplicate the imported records over columns 0 and 5. If a string like `"1-6"` is used as argument, the deduplication will use the 2nd through 7th fields. If an error occurs, the method will return `Nothing`.

### ☕Example

```vb
Sub DedupeCSV()
    Dim CSVint As CSVinterface
    Dim DedupedData As CSVArrayList
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        Set DedupedData = .Dedupe("5-8,11")        'Deduplicate using fields indexes 5 through 8 and 11.
        Set DedupedData = .Dedupe("1,5,6")         'Deduplicate using fields indexes 1, 5 and 6.
    End With
    Set CSVint = Nothing
    Set DedupedData = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)