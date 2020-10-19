---
title: ImportFromCSV
parent: Methods
grand_parent: API
nav_order: 3
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

---

## Remarks

If the *HeadersOmission* parameter is set to `True`, the CSV file headers, first line, will be ignored by the parser only when the `StartingRecord` property is set to 1. 

See also:
 [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html),
 [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html),
 [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html),
 [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html),
 [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html).

---

## Behavior

The `FieldsDelimiter`, `RecordsDelimiter`, `EscapeChar`, `StartingRecord` and `EndingRecord` properties sets the The `ImportFromCSV` method behavior. If the CSV file already exist on path, the `ImportFromCSV` method will overwrites all its content. If that is not the case, a new file will be created.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)