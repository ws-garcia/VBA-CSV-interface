---
title: SortByField
parent: Methods
grand_parent: API
nav_order: 28
---

# SortByField
{: .fs-6 }

Sorts the imported CSV/TSV data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`SortByField`*(\[fromIndex:= -1\], \[toIndex:= -1\], \[SortingKey:= 1\], \[SortAlgorithm:= SortingAlgorithms.SA_Quicksort\])*

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
<td style="text-align: left;"><em>SortingKey</em></td>
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
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html), [Sort method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/sort.html).

---

## Behavior

When the *fromIndex* parameter is omitted, the sorting of the data starts at the first field of the record specified in the *SortingKey* parameter. Omitting the *toIndex* parameter takes the sorting to the last available field. If the *SortingKey* parameter is omitted, the data will be sorted on the first field in ascending order, set this parameter to a negative `Integer` to sort the data in descending order on the given record (e.g. `SortingKey:=-2` will sort in descending order on the second record). In addition, the user can pass a one-dimensional array in the *SortingKey* parameter to achieve multilevel data sorting on several fields at once.

### ☕Example

```vb
Sub SortByField()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        .SortByField SortingKey:=1, SortAlgorithm:=SA_Quicksort                 'Sort the data in ascending order on header record.
                                                                                'The operation will change fields order instead
                                                                                'of records ordering.
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
