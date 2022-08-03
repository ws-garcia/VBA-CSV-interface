---
title: SplitField
parent: Methods
grand_parent: API
nav_order: 30
---

# SplitField
{: .fs-6 }

Splits the specified field in the imported CSV data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`SplitField`*(aIndex, CharToSplitWith)*

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
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable representing the index of the field to be splited.</td>
</tr>
<tr>
<td style="text-align: left;"><em>CharToSplitWith</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Represents the character to be used in the split operation.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

## Behavior

The `SplitField` method will split the field specified in the current instance, if all records have the same number of fields, using the character specified via `CharToSplitWith`. 

### ☕Example

```vb
Sub SplitField()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_file.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        On Error Resume Next
        .SplitField 1, "|"											'Split field at index 1 using a pipe character.
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
