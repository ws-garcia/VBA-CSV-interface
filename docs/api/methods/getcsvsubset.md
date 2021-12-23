---
title: GetCSVsubset
parent: Methods
grand_parent: API
nav_order: 11
---

# GetCSVsubset
{: .fs-6 }

Returns a set of records matching the criteria applied to a desired field from a CSV file.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`GetCSVsubset`*(filePath, filters, keyIndex, \[configObj:= Nothing\])*

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
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable representing the full path to the target CSV file.</td>
</tr>
<tr>
<td style="text-align: left;"><em>filters</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Variant</code> Type variable representing an array containing all the criteria to be applied to the desired field.</td>
</tr>
<tr>
<td style="text-align: left;"><em>keyIndex</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable representing the index of the field to apply the criteria.</td>
</tr>
<tr>
<td style="text-align: left;"><em>configObj</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>CSVparserConfig</code> object variable holding the parser configuration.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVArrayList` object

See also
: [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html).

---

## Behavior

The `GetCSVsubset` method will retrieve all records where the field at position *keyIndex* meets all given criteria. If *configObj* is not given, the internal configuration of the current instance will be used.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>If the *filters* parameter is not a two dimensional array, an error will occur.
{: .text-grey-dk-300 .bg-yellow-000 }

### ☕Example

```vb
Private Sub GetCSVSubSet()
    Dim CSVint As CSVinterface
    Dim CSVrecords As CSVArrayList
    Dim path As String
    Dim conditions() As String
    Dim queryFilters As Variant
    
    Set CSVint = New CSVinterface
    path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    CSVint.parseConfig.Headers = False                                      'The file has no header record/row
    ReDim conditions(0 To 1)
    conditions(0) = "Asia": conditions(1) = "Europe"
    queryFilters = conditions
    If path <> vbNullString Then
        Set CSVrecords = CSVint.GetCSVSubSet(path, queryFilters, 1)         'Data filtered on first field
        CSVint.DumpToSheet DataSource:=CSVrecords                           'Dump result to a new sheet
        Set CSVint = Nothing
        Set CSVrecords = Nothing
    End If
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)