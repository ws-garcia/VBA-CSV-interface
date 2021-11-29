---
title: InsertRecord
parent: Methods
grand_parent: API
nav_order: 17
---

# InsertRecord
{: .d-inline-block }

New
{: .label .label-purple }

Inserts a new record, at the specified position, in the imported CSV data.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`InsertRecord`*(aIndex)*

### Parameters

The required *aIndex* argument is an identifier specifying a `Long` Type variable. Represents the index in which the new record will be inserted.

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `InsertRecord` method will insert an empty record, if all imported records have the same number of fields, into the imported data in the current instance.

### ☕Example

```vb
Sub InsertRecord()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Chinese CSV.csv"
        .utf8EncodedFile = True                                         'The file is UTF-8 encoded
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        .InsertRecord 5                                                 'Insert a record"
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)