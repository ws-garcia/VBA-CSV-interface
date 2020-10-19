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

*expression*.`ExportToCSV` *(csvArray)*

### Parameters

The required *csvArray* argument is an identifier specifying a `Variant` array variable.

### Return value

_None_

---

## Remarks

The *csvArray* parameter must be declared as `Variant` array. Passing a variable that isn't an array will cause an error and the operation aborts. 

Before invoke the method, the user must to set the the `FieldsDelimiter`, `RecordsDelimiter` and the `EscapeChar` properties in order to fit the class behavior to the needs.

See also:
 [FieldsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html),
 [RecordsDelimiter property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/recordsdelimiter.html),
 [EscapeChar property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html).

---

## Behavior

If the CSV file already exist on path, the `ExportToCSV` method will overwrites all its content. If that is not the case, a new file will created.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)