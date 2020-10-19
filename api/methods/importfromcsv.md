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

**Note: Before invoke the `ImportFromCSV` method, the user must to open a connection to the CSV file.**

If the *HeadersOmission* parameter is set to `True`, the CSV file headers, first line, will be ignored by the parser only when the `StartingRecord` property is set to 1. 

See also
: [OpenConnection method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openconnection.html), [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html), [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html), [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html), [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html).

---

## Behavior

The `FieldsDelimiter`, `RecordsDelimiter`, `EscapeChar`, `StartingRecord` and `EndingRecord` properties sets the method's behavior to the needs.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)