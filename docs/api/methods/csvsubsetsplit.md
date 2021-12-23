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

*expression*.`CSVsubsetSplit`*(filePath, \[subsetColumn:= 1\], \[headers:= True\])*

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
<td style="text-align: left;"><em>subsetColumn</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable representing the index of the field on which the creation of the data groups will take place.</td>
</tr>
<tr>
<td style="text-align: left;"><em>headers</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable indicating whether the target CSV file has a header record.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `Collection` object

---

## Behavior

The `CSVsubsetSplit` method will create a file for each different value (data grouping) in the field at the *subsetColumn* position, then all related data is appended to the respective file. Use the *headers* parameter to include a header record in each new CSV file. When the CSV file has a header record and the user sets the *header* parameter to `False`, the header row is saved in a separate file and the rest of CSV files will have no header record.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The result subsets will be saved in a folder named [\*-subsets], where (\*) denotes the name of the source CSV file.
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
