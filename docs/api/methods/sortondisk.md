---
title: SortOnDisk
parent: Methods
grand_parent: API
nav_order: 29
---

# SortOnDisk
{: .d-inline-block }

New
{: .label .label-purple }

Sorts a CSV file on disk rather than in memory.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`SortOnDisk`*(filePath, \[SortingKeys:= 1\], \[Headers:= True\], \[streamSize := 20\], \[SortAlgorithm:= SortingAlgorithms.SA_Quicksort\], \[ExportationBunchSize:= 10000\])*

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
<td style="text-align: left;"><em>filePath</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable representing the CSV file path, including file extension.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SortingKeys</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Variant</code> Type variable representing the columns/keys for the logical comparisons.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Headers</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>streamSize</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Single</code> Type variable representing the single textstream factor size.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SortAlgorithm</em></td>
<td style="text-align: left;">Optional. Identifier specifying a member of the <code>SortingAlgorithms</code> Enumeration.</td>
</tr>
<tr>
<td style="text-align: left;"><em>ExportationBunchSize</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable representing the amount of items to export in a single operation.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `String`

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The sorting operation will require intensive I/O usage of the hard disk drive, so the performance of the method will also be tied to the R/W speed of the disk.
{: .text-grey-dk-300 .bg-yellow-000 }

See also
: [CSVTextStream class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvtextstream.html), [Sort method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/sort.html).

---

## Behavior

The data sorting start on the second record when the *Headers* parameter is set to `True`. If the *SortingKeys* parameter is omitted the data will be sorted on the first column/field in ascending order, set this parameter to a negative `Integer` to sort the data in descending order on the given column (e.g. `SortingKeys:=-2` will sort in descending order on the second field). In addition, the user can pass a one-dimensional array in the *SortingKeys* parameter to achieve multilevel data sorting on several fields at once. The returned string is the full path to the new sorted file whose name has the form "\*-sorted.csv" where "\*" represents the name of the CSV file to be sorted.


### ☕Example

```vb
Sub SortOnDisk()
    Dim CSVint As CSVinterface
    Dim SortKeys() As Long
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    ReDim SortKeys(0 To 2)
    SortKeys(0) = -1: SortKeys(1) = 5: SortKeys(2) = -11
    With CSVint
        .SortOnDisk .parseConfig.path, sortingKeys:=SortKeys    'Sort the data in descending order on column 1,
                                                                'then sort in ascending order on column 5 and
                                                                'sort in descending order on column 11. This
                                                                'multi-level sorting is "stable".
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
