---
title: ImportFromCSV
parent: Methods
grand_parent: API
nav_order: 5
---

# ImportFromCSV
{: .fs-9 }

Imports a CSV file's content to an array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ImportFromCSV`*({HeadersOmission:= `False`}, {PassControlToOS:= `True`})*

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
<td style="text-align: left;"><em>HeadersOmission</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>PassControlToOS</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> variable.</td>
</tr>
</tbody>
</table>

### Return value

_None_

>📝**Note:**
>If the *HeadersOmission* parameter is set to `True`, the CSV file headers, first record, will be ignored by the parser only when the `StartingRecord` property is set to 1. 
>The *PassControlToOS* parameter allows user to pass control to the operating system. Control is returned after the operating system has finished processing the events in its queue.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html), [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html), [CommentLineIndicator property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/commentlineindicator.html).

---

## Behavior

User can set `CommentLineIndicator` for those CSV files having a combination of empties lines, blanks lines or commented ones for parse the file ONLY when the parser is working on `QuotationMode.Critical` mode. In that mode, the cited lines are simply skipped, leaving no empty values between records separated by this kind of lines. In other words, if the CSV file holds a record and then some special lines (blank, empty or commented) and then another record, the second record will be saved contiguous to the first record ignoring the lines between both.

>⚠️**Caution:**
>Before invoke the `ImportFromCSV` method, the user must to open a connection to the CSV file. If the CSV file has no data, this is the file is an empty one, the `ImportFromCSV` method returns an empty array, that is, an array bounded from 0 to -1 and holding no elements and no data.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
