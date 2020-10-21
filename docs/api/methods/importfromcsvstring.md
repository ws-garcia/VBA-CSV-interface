---
title: ImportFromCSVstring
parent: Methods
grand_parent: API
nav_order: 5
---

# ImportFromCSVstring
{: .fs-9 }

Parses a string and save its CSV data to an array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ImportFromCSVstring`*(CSVstring, {HeadersOmission})*

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
<td style="text-align: left;"><em>CSVstring</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> variable to be parsed.</td>
</tr>
<tr>
<td style="text-align: left;"><em>HeadersOmission</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> variable.</td>
</tr>
</tbody>
</table>

### Return value

_None_

---

## Remarks

If the *HeadersOmission* parameter is set to `True`, the *CSVstring* headers, first record, will be ignored by the parser only when the `StartingRecord` property is set to 1. 

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html), [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html).

---

## Behavior

The `FieldsDelimiter`, `RecordsDelimiter`, `EscapeChar`, `StartingRecord` and `EndingRecord` properties sets the method's behavior to the needs.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
