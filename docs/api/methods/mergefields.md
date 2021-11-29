---
title: MergeFields
parent: Methods
grand_parent: API
nav_order: 18
---

# MergeFields
{: .d-inline-block }

New
{: .label .label-purple }

Merges the specified fields in the imported CSV data.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`MergeFields`*(indexes, CharToMergeWith)*

### Parameters

<table>
<thead>
<tr>
<th style="text-align: left;">Part</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>indexes</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable representing the indexes to be merged.</td>
</tr>
<tr>
<td style="text-align: left;"><em>CharToMergeWith</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Represents the character to be used in the merge operation.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `MergeFields` method will merge all fields specified in the current instance, if all records have the same number of fields, using the character specified via `CharToMergeWith`. 

The `indexes` parameter will indicate which fields/columns will be merged. A string like `"2,7"` used as parameter will merge the imported records over the columns with indexes 2 and 7. If a string like `"3-8,10"` is used as argument, the merge operation will use the 4th to 9th fields and the 11th field. 

### ☕Example

```vb
Sub MergeFields()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        On Error Resume Next
        .MergeFields "0-3,11", "|"                                          'Merge fields at indexes 0 to 3 and 11 using a pipe character.
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)