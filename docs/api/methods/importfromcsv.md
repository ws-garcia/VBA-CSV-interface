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

*expression*.`ImportFromCSV`*({HeadersOmission})*

### Parameters

The optional *HeadersOmission* argument is an identifier specifying a `Boolean` variable.

### Return value

_None_

>:pencil: **NOTE:**
>
> Before invoke the `ImportFromCSV` method, the user must to open a connection to the CSV file. If the *HeadersOmission* parameter is set to `True`, the CSV file headers, first record, will be ignored by the parser only when the `StartingRecord` property is set to 1. 

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html), [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html), [CommentLineIndicator property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/commentlineindicator.html).

---

## Behavior

User can set `CommentLineIndicator` for those CSV files having a combination of empties lines, blanks lines or commented ones for parse the file ONLY when the parser is working on `QuotationMode.Critical` mode. In that mode, the cited lines are simply skipped, leaving no empty values between records separated by this kind of lines. In other words, if the CSV file holds a record and then some special lines (blank, empty or commented) and then another record, the second record will be saved contiguous to the first record ignoring the lines between both.

>:warning: **CAUTION**
>
>If the CSV file has no data, this is the file is an empty one, the `ImportFromCSV` method returns an empty array, that is, an array bounded from 0 to -1 and holding no elements and no data.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
