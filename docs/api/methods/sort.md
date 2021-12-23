---
title: Sort
parent: Methods
grand_parent: API
nav_order: 27
---

# Sort
{: .fs-6 }

Sorts the imported CSV/TSV data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`Sort`*(\[fromIndex:= -1\], \[toIndex:= -1\], \[SortingKeys:= 1\], \[SortAlgorithm:= SortingAlgorithms.SA_IntroSort\])*

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
<td style="text-align: left;"><em>fromIndex</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>toIndex</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SortingKeys</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Variant</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SortAlgorithm</em></td>
<td style="text-align: left;">Optional. Identifier specifying a member of the <code>SortingAlgorithms</code> Enumeration.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before sort data, is required to make a call to the `ImportFromCSV` or `ImportFromCSVstring` method.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html), [SortByField method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/sortbyfield.html).

---

## Behavior

When the *fromIndex* parameter is omitted the data sorting start on the second record when the `ParseConfig.headers` is set to `True` and `ParseConfig.headersOmission` is set to `False`. Omitting the *toIndex* parameter takes the sorting to the last available record. If the *SortingKeys* parameter is omitted the data will be sorted on the first column/field in ascending order, set this parameter to a negative `Integer` to sort the data in descending order on the given column (e.g. `SortingKeys:=-2` will sort in descending order on the second field). In addition, the user can pass a one-dimensional array in the *SortingKeys* parameter to achieve multilevel data sorting on several fields at once.

### ☕Example

```vb
Sub Sort()
    Dim CSVint As CSVinterface
    Dim SortKeys() As Long
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    ReDim SortKeys(0 To 2)
    SortKeys(0) = -1: SortKeys(1) = 5: SortKeys(2) = -11
    With CSVint
        .ImportFromCSV .parseConfig
        .Sort SortingKeys:=SortKeys, SortAlgorithm:=SA_Quicksort                'Sort the data in descending order on column 1,
                                                                                'then sort in ascending order on column 5 and
                                                                                'sort in descending order on column 11. This
                                                                                'multi-level is "stable".
        .DumpToSheet
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
