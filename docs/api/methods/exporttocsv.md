---
title: ExportToCSV
parent: Methods
grand_parent: API
nav_order: 4
---

# ExportToCSV
{: .fs-9 }

Exports an array's content to a CSV file.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ExportToCSV`*(csvArray, {PassControlToOS:= `True`})*

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
<td style="text-align: left;">Required. Identifier specifying a <code>Variant</code> array variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>PassControlToOS</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> variable.</td>
</tr>
</tbody>
</table>

### Return value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before invoke the `ExportToCSV` method, the user must to open a connection to the CSV file. The *csvArray* parameter must be declared as `Variant` array. Passing a variable that isn't an array will cause an error and the operation aborts. 
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html).

---

## Behavior

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
Use jagged arrays keeping in mind they can make the VBA hosting application run out of memory due the `Variant` data Type is a memory hog.
{: .text-grey-dk-300 .bg-yellow-000 }

The `FieldsDelimiter`, `RecordsDelimiter` and `EscapeChar` properties sets the method's behavior to the needs.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)