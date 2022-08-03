---
title: ShiftRecord
parent: Methods
grand_parent: API
nav_order: 25
---

# ShiftRecord
{: .fs-6 }

Moves the Record the specified number of times in the current instance.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`ShiftRecord`*(aIndex, Shift)*

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
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the index of the Record to be shifted.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Shift</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the number of positions the Record will shift.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `ShiftRecord` method will move the record down if the `Shift` is a positive integer, otherwise the record will be shifted up. 

### ☕Example

```vb
Sub ShiftRecord()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        On Error Resume Next
        .ShiftRecord 4, -1                              'Shift the 5th record up by 1 position
        .ShiftRecord 3, 1                               'Shift the 4th record down by 1 position
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)