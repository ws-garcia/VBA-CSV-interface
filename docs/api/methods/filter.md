---
title: Filter
parent: Methods
grand_parent: API
nav_order: 10
---

# Filter
{: .d-inline-block }

New
{: .label .label-purple }

Returns a list of records as a result of applying filters on the target CSV file or imported CSV data using expression evaluation.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`Filter`*(Pattern, [FilePath], [ExcludeFirstRecord])*

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
<td style="text-align: left;"><em>Pattern</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Represents a valid string expression to evaluate when filtering records</td>
</tr>
<tr>
<td style="text-align: left;"><em>FilePath</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> Type variable. Represents the full file path, including file extension, of the CSV file used for data filtering.</td>
</tr>
<tr>
<td style="text-align: left;"><em>ExcludeFirstRecord</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable. When <code>True</code>, the file headers will be excluded.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVArrayList`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html).

## Behavior

The `Pattern` parameter is evaluated according to the number of records in the CSV file, when the evaluation returns `True`, the current record is saved. The rules that apply to the `Pattern` parameter are listed below:
* To reference a field value, the user must type something like `f#` where `f` is a required identifier and `#` is the numeric position of the desired field. For example, `f1>5` indicates the selection of records whose first field value is greater than `5`.
* If the user needs to compare literal strings, the values must be enclosed in apostrophes. Example, `Region = 'Central America'` is a valid string assigned to the variable `Region`.
* User can use functions in the `Pattern` definition, including custom UDFs (refer to [VBAexpressions documentation](https://github.com/ws-garcia/VBA-Expressions)). I.e.: `min(f5;f2)>=100` 

When the `FilePath` argument is omitted, the method will proceed to filter the data stored in the current instance, otherwise it will filter the content of the CSV file specified with the referred argument.

### ☕Example

```vb
Sub FilterCSV()
    Dim CSVint As CSVinterface
    Dim path As String
    Dim FilteredData As CSVArrayList
    
    Set CSVint = New CSVinterface
    path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    CSVint.parseConfig.Headers = False                                      		'The file has no header record/row
    CSVint.parseConfig.path = path
    If path <> vbNullString Then
        Set FilteredData = CSVint.Filter("f1='Asia' & f9>20 & f9<=50", path) 		'Select "Units sold" greater than 20 and less or 
																											'equal to 50 from Asian customers
        Set CSVint = Nothing
        Set FilteredData = Nothing
    End If
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)