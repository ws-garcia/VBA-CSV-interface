---
title: ExportToCSV
parent: Methods
grand_parent: API
nav_order: 9
---

# ExportToCSV
{: .fs-9 }

Exports an array's content to a CSV/TSV file.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ExportToCSV`*(csvArray, \[pconfig:= `Nothing`\], \[PassControlToOS:= `True`\], \[enableDelimiterGuessing:= `True`\], \[EnforcedQuotation:= `False`\])*

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
<td style="text-align: left;"><em>csvArray</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Variant</code> Type variable which represents the data to be written to a CSV file.</td>
</tr>
<tr>
<td style="text-align: left;"><em>pconfig</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>CSVparserConfig</code> object variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>PassControlToOS</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable. Use this to prevent the application from appearing to hang at runtime.</td>
</tr>
<tr>
<td style="text-align: left;"><em>enableDelimitersSniffing</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>EnforcedQuotation</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
</tbody>
</table>

### Returns value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The *csvArray* parameter can be an `CSVArrayList` or an array (one-dimensional, two-dimensional or jagged) variable, passing another type of variable will cause an error. 
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html), [CSVArrayList class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html).

---

## Behavior

When the `pconfig` parameter is omitted, the parser will use the `ParseConfig` property as the configuration object. If the file specified in the `.path` configuration property already exists and has some content, the parser will try to guess the delimiters and the data will be added to the file when the `enableDelimitersSniffing` is set to `True`. Setting the `EnforcedQuotation` parameter to `True` will force to quote all fields in the created CSV file. 

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>Fields containing literal escape characters will be escaped using the classic escape sequence (duplicating each escape character) or using the Unix escape sequence. The behavior will be controlled by the `.dialect.escapeMode` property of the given configuration object.
{: .text-grey-dk-300 .bg-yellow-000 }

### ☕Example

```vb
Sub ExportToTSV()
    Dim CSVint As CSVinterface
    Dim conf As CSVparserConfig

    Set CSVint = New CSVinterface
    Set conf = CSVint.parseConfig
    With conf
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
        .dynamicTyping = False
        Set .dialect = CSVint.SniffDelimiters(conf)   								'Try to guess CSV file data delimiters
        CSVint.ImportFromCSV conf 															'Import the data
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.tsv"
        .dialect.fieldsDelimiter = vbTab
        CSVint.ExportToCSV CSVint.items, conf 											'Export internal items to a TSV file
    End With
    Set CSVint = Nothing 																		'Terminate the current instance
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)