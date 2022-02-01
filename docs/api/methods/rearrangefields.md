---
title: RearrangeFields
parent: Methods
grand_parent: API
nav_order: 20
---

# RearrangeFields
{: .d-inline-block }

New
{: .label .label-purple }

Rearranges the fields of the imported CSV data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`RearrangeFields`*(FieldsOrder)*

### Parameters

The required *FieldsOrder* argument is an identifier specifying a `String` Type variable. Represents the new order of all fields.

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `RearrangeFields` method will change the order of all fields specified in the current instance, if all records have the same number of fields. The `FieldsOrder` parameter will indicate the new order of the fields/columns. The method requires specifying a position for all sigle fields. A string such as `"0-3,5-4,6-11"` used as a parameter will leave the position of fields with indexes 0 to 3 unchanged, swap the fields at indexes 5 and 4, and leave all remaining fields in position. 

### ☕Example

```vb
Sub Rearrange()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        On Error Resume Next
        .RearrangeFields "0-7,10-8,11"                                      'Leave unchanged fields at indexes from
                                                                            '0 to 7, swap the field at index 8 and 10.
                                                                            'Field at index 9 and 11 remain in its
                                                                            'position.
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
