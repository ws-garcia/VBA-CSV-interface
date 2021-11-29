---
title: ImportFromCSV
parent: Methods
grand_parent: API
nav_order: 14
---

# ImportFromCSV
{: .fs-9 }

Imports a CSV/TSV file's content to the current instance.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ImportFromCSV`*(configObj, \[FilterColumns\])*

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
<td style="text-align: left;"><em>configObj</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>CSVparserConfig</code> object variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>FilterColumns</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>ParamArray</code> of <code>Variant</code> Type variable.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html), [CSVTextStream class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvtextstream.html).

---

## Behavior

The `configObj` parameter is an object with all the options considered by the parser during the import operation, see the [ParseConfig Property documentation](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html). User can use the `FilterColumns` parameter for retrieve only certain fields from each CSV/TSV record. The filters can be strings representing the names of the fields determined with the header record, or numbers representing the position of the requested field. If not filters defined, all the fields of the requested records will be retrieved.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>If the target file has no data (the file is an empty one) or an error occur when parsing, the `ImportFromCSV` method returns a non-initialized object.
{: .text-grey-dk-300 .bg-yellow-000 }

### ☕Example

```vb
Sub ImportFromCSV()
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
        Set .dialect = CSVint.SniffDelimiters(CSVint.parseConfig)               'Sniff delimiters and save the result in the config object
    End With
    With CSVint
        .ImportFromCSV .parseConfig                                             'Import CSV data
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)