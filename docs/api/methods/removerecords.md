---
title: RemoveRecords
parent: Methods
grand_parent: API
nav_order: 22
---

# RemoveRecords
{: .d-inline-block }

New
{: .label .label-purple }

Removes a field, at the specified position, from the imported CSV data.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`RemoveRecords`*(aIndex, \[count\])*

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
<td style="text-align: left;"><em>aIndex</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the index at which the records will be deleted.</td>
</tr>
<tr>
<td style="text-align: left;"><em>count</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable. Represents the number of records to be deleted.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `RemoveRecords` method will remove an specified amount of records from the current instance, if they all have the same number of fields. 

### ☕Example

```vb
Sub RemoveRecord()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Chinese CSV.csv"
        .utf8EncodedFile = True
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        .RemoveRecords 2, 2                              'Remove 2 records starting at the 3th record
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)