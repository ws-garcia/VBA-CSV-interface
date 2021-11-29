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

Returns a list of records as a result of applying filters on the imported CSV data.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`Filter`*(fieldIndex, patterns)*

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
<td style="text-align: left;"><em>fieldIndex</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Represents the index of the field used for data filtering.</td>
</tr>
<tr>
<td style="text-align: left;"><em>patterns</em></td>
<td style="text-align: left;">Optional. Identifier specifying a list of <code>Strings</code> Type variables.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVArrayList`

---

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html).

## Behavior

If the `patterns` parameter is omitted the complete set of stored data will be returned. The rules that apply to the `patterns` parameter are listed below:
* The comparison is influenced by the `Option Compare` statement (one of: `Option Compare Binary` or `Option Compare Text`). The binary compare is case sensitive, the text compare is not.
* The following table describes the special characters to be used when creating patterns; all other characters match themselves:
	* 
	|Character|Meaning|
	|:------:|:-----|
	|?|Any single character|
	|\*|Zero or more characters|
	|#|Any single digit (0-9)|
	|\[list\]|Any single character in list|
	|\[!list\]|Any single character not in list|
	|\[\]|A zero-length string ("")|
* 'list' matches a group of characters in `patterns` to a single character in the string and can contain almost all available characters, including digits.
* Use a hyphen (-) in 'list' to create a range of characters that matches a character in the string: e.g. [A-D] matches A,B,C, or D at that character position in the string. Multiple ranges of characters can be included in 'list' without the use of a delimiter: e.g. \[A-DJ-L\].
* Use the hyphen at the start or end of 'list' to match to itself. For example, \[-A-G\] matches a hyphen or any character from A to G.
* The exclamation mark in the "pattern" match is similar to the negation operator. For example, [!A-G] matches all characters except characters A through G.
* The exclamation mark outside the bracket  matches itself.
* To use any special character as a matching character, enclose the special character in brackets. For example, to match a question mark, use \[?\].

### ☕Example

```vb
Sub FilterCSV()
    Dim CSVint As CSVinterface
    Dim FilteredData As CSVArrayList
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    End With
    With CSVint
        .ImportFromCSV .parseConfig
        Set FilteredData = .Filter(11, "###.##", "####.##")   'Filter data between hundreds and thousands
    End With
    Set CSVint = Nothing
    Set FilteredData = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)