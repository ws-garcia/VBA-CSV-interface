---
title: OpenConnection
parent: Methods
grand_parent: API
nav_order: 4
---

# OpenConnection
{: .fs-9 }

Loads a CSV file on memory for data Input/Output operations.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`OpenConnection`*(csvPathAndFilename)*

### Parameters

The required *csvPathAndFilename* argument is an identifier specifying a `String` variable.

### Return value

_None_

---

## Remarks

The `OpenConnection` method is the preamble to the `ImportFromCSV` and `ExportToCSV` methods, this means each call to the citated methods must be preceded by a `OpenConnection` method call.

After call the `OpenConnection` method is possible to check if the instance is bind to the CSV file, for which is only needed to read the current instance `Connected` property.

See also
: [Connected property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/connected.html), [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ExportToCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/exporttocsv.html).

---

## Behavior

When the given path exists the file will be created on that path, otherwise an error occur. For on path existing CSV file, the `OpenConnection` method will overwrites all its content. If that is not the case, a new file will be created.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)