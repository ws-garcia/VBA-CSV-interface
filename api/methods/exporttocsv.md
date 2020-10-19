---
title: ExportToCSV
parent: Methods
grand_parent: API
nav_order: 2
---

# ExportToCSV
{: .fs-9 }

Exports an array's content to a CSV file.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ExportToCSV`*(csvArray)*

### Parameters

The required *csvArray* argument is an identifier specifying a `Variant` array variable.

### Return value

_None_

---

## Remarks

**Note: Before invoke the `ExportToCSV` method, the user must to open a connection to the CSV file.**

The *csvArray* parameter must be declared as `Variant` array. Passing a variable that isn't an array will cause an error and the operation aborts. 

See also:
 [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html),
 [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html),
 [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html).

---

## Behavior

The `FieldsDelimiter`, `RecordsDelimiter` and `EscapeChar` properties sets the method's behavior to the needs.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)