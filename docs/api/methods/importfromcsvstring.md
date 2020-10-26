---
title: ImportFromCSVstring
parent: Methods
grand_parent: API
nav_order: 6
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

**Note: Before invoke the `ImportFromCSV` method, the user must to open a connection to the CSV file.**

If the *HeadersOmission* parameter is set to `True`, the CSV file headers, first record, will be ignored by the parser only when the `StartingRecord` property is set to 1. 

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html), [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html), [CommentLineIndicator property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/commentlineindicator.html).

---

## Behavior

User can set `CommentLineIndicator` for such CSV files having a combination of empties lines, blanks lines or commented ones can be parsed only if the parser is working on `QuotationMode.Critical` mode. In that mode, the cited lines are simply skipped, leaving no empty values between records separated by this kind of lines.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)