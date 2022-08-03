---
title: RemoveField
parent: Methods
grand_parent: API
nav_order: 21
---

# RemoveField
{: .fs-6 }

Removes a field, at the specified position, from the imported CSV data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`RemoveField`*(aIndex)*

### Parameters

The required *aIndex* argument is an identifier specifying a `Long` Type variable.  Represents the index in which the field will be removed.

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `RemoveField` method will remove a field from all records, if they all have the same number of fields, from the current instance.

### ☕Example

```vb
Sub RemoveField()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Chinese CSV.csv"
        .utf8EncodedFile = True                                         'The file is UTF-8 encoded
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        .RemoveField 0                                                  'Remove the first field from all records"
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)