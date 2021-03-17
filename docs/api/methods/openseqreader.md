---
title: OpenSeqReader
parent: Methods
grand_parent: API
nav_order: 16
---

# OpenSeqReader
{: .d-inline-block }

New
{: .label .label-purple }

Opens a sequential CSV reader for import records one at a time.
{: .fs-6 .fw-300 }

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
<td style="text-align: left;">Required. Identifier specifying a <code>ParserConfig</code> object variable.</td>
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
: [GetRecord Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/getrecord.html), [CloseSeqReader Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/closeseqreader.html).

---

## Behavior

The `OpenSeqReader` method works in conjunction with the `GetRecord` method. The `FilterColumns` parameter is used to retrieve only certain fields from each CSV/TSV record. Filters can be strings representing the names of the fields determined with the header record, or numbers representing the position of the requested field. If no filters are defined, all fields of the requested records will be retrieved. Each call to the `OpenSeqReader` method will create a new conection to the CSV file.


>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>After opening a sequential reader, the user can read the CSV records one by one and implement logics to work with the extracted data. This makes it possible to mimic the more complex behavior of SQL statements.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)