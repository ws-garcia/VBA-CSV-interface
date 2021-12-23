---
title: ShiftField
parent: Methods
grand_parent: API
nav_order: 24
---

# ShiftField
{: .d-inline-block }

New
{: .label .label-purple }

Moves the field the specified number of times in all records of the current instance.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`ShiftField`*(aIndex, Shift)*

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
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the index of the field to be shifted.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Shift</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the number of positions the field will shift.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `ShiftField` method will move the field to the right if the `Shift` is a positive integer, otherwise the field will be shifted to the left. 

### ☕Example

```vb
Sub ShiftField()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        On Error Resume Next
        .ShiftField 1, 1                          'Shift the 2nd field to the right by one position
        .ShiftField 2, -1                         'Shift the 3th field to the left by one position
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)