---
title: SniffDelimiters
parent: Methods
grand_parent: API
nav_order: 26
---

# SniffDelimiters
{: .fs-6 }

Returns a CSV dialect after run an analysis over a String variable or in the CSV/TSV file indicated in the `.path` property of the configuration object.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`SniffDelimiters`*(confObject, \[CSVstring\] = vbNullString)*

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
<td style="text-align: left;"><em>confObject</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>CSVparserConfig</code> object variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>CSVstring</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> Type variable.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVdialect`

---

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html), [CSVSniffer class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvsniffer.html).

## Behavior

The parser will guess the delimiters in the CSV file only when the `CSVstring` parameter is set to `vbNullString`, otherwise the guessing occurs on the given string.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>Only some records will be used to guess the delimiters. The method is very accurate, but there is a risk of inaccuracy in some rare cases.
{: .text-grey-dk-300 .bg-yellow-000 }

### ☕Example

```vb
Sub SniffInString()
    Dim CSVint As CSVinterface
    Dim CSVdata As String
    Dim fPath As String
    
    Set CSVint = New CSVinterface
    fPath = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
    With CSVint
        CSVdata = .GetDataFromCSV(fPath)
        Set .parseConfig.dialect = .SniffDelimiters(.parseConfig, CSVdata)      'Sniff delimiters and save to config object
    End With
    Set CSVint = Nothing
End Sub
Private Sub SniffInFile()
    Dim CSVint As CSVinterface
    Dim csvRecord As CSVArrayList
            
    Set CSVint = New CSVinterface
    With CSVint
        .parseConfig.path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
        Set .parseConfig.dialect = .SniffDelimiters(.parseConfig)                           'Sniff delimiters and save to config object
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)