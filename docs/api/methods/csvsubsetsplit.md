---
title: CSVsubsetSplit
parent: Methods
grand_parent: API
nav_order: 3
---

# CSVsubsetSplit
{: .fs-6 }

Splits the CSV data into a set of files in which each piece has a related portion of the data.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`CSVsubsetSplit`*(filePath, \[subsetColumns:= 1\], \[headers:= True\], \[repeatHeaders:= True\], \[streamSize:= 20\], \[oConfig:=Nothing\])*

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
<td style="text-align: left;"><em>subsetColumns</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Variant</code> Type variable representing the indexes of the fields on which the data groups will be created.</td>
</tr>
<tr>
<td style="text-align: left;"><em>headers</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable indicating whether the target CSV file has a header record.</td>
</tr>
<tr>
<td style="text-align: left;"><em>repeatHeaders</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable indicating whether the header record, from the target CSV file, will be copied to all created files.</td>
</tr>
<tr>
<td style="text-align: left;"><em>streamSize</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable representing the buffer size factor used to read the target CSV file.</td>
</tr>
<tr>
<td style="text-align: left;"><em>oConfig</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>CSVparserConfig</code> Object variable holding all the configurations to parse the CSV file.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `Collection` object

---

## Behavior

The `CSVsubsetSplit` method will create a file for each different value (data grouping) in the fields at the *subsetColumns* position, then all related data is appended to the respective file. Use the *headers* parameter to include a header record in each new CSV file. The *subsetColumns* parameter can be a single value or an array of `Long` values.  When the CSV file has a header record and the user sets the *header* parameter to `False`, the header row is saved in a separate file and the rest of CSV files will have no header record. The user can control when to include the headers by using the *repeatHeaders* parameter.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The result subsets will be saved in a folder named [\*-WorkDir], where (\*) denotes the name of the source CSV file.
{: .text-grey-dk-300 .bg-grey-lt-000 }

### ☕Example

```vb
Sub SplitCSV()
    Dim CSVint As CSVinterface
    Dim path As String
    
    Set CSVint = New CSVinterface
    path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    CSVint.CSVsubsetSplit path, 3, True   ' Split the CSV and rank the resulting files by
                                          ' the contents of the third column. Header is
                                          ' assumed to be present on the file.
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
