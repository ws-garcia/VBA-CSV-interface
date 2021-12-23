---
title: OpenSeqReader
parent: Methods
grand_parent: API
nav_order: 19
---

# OpenSeqReader
{: .fs-6 }

Opens a sequential CSV reader for import records one at a time.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`OpenSeqReader`*(configObj, \[FilterColumns\])*

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

_None_

See also
: [GetRecord Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/getrecord.html), [CloseSeqReader Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/closeseqreader.html), [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html), [CSVTextStream class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvtextstream.html).

---

## Behavior

The `OpenSeqReader` method works in conjunction with the `GetRecord` method. The `configObj` parameter is an object with all the options considered by the parser during the import operation, see the [ParseConfig Property documentation](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html). The `FilterColumns` parameter is used to retrieve only certain fields from each CSV/TSV record. Filters can be strings representing the names of the fields determined with the header record, or numbers representing the position of the requested field. If no filters are defined, all fields of the requested records will be retrieved. Each call to the `OpenSeqReader` method will create a new conection to the CSV file.


>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>After opening a sequential reader, the user can read the CSV records one by one and implement logics to work with the extracted data. This makes it possible to mimic the more complex behavior of SQL statements.
{: .text-grey-dk-300 .bg-grey-lt-000 }

### ☕Example

```vb
Private Sub OpenSeqReaderAndGetRecord()
    Dim CSVint As CSVinterface
    Dim csvRecord As CSVArrayList
            
    Set CSVint = New CSVinterface
    With CSVint
        .parseConfig.path = Environ("USERPROFILE") & "\Desktop\Demo_100000records.csv"
        Set .parseConfig.dialect = .SniffDelimiters(.parseConfig)
        .OpenSeqReader .parseConfig, 1, 2                                                  'Start sequential reader
                                                                                           'Import only 1st and 2nd fields
        Do
            Set csvRecord = .GetRecord                                                      'Get a record from CSV file
        Loop While Not csvRecord Is Nothing                                                 'Loop trhonght all records in file
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)