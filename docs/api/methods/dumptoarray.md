---
title: DumpToArray
parent: Methods
grand_parent: API
nav_order: 1
---

# DumpToArray
{: .fs-9 }

Dumps the data from the current instance to an array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToArray`*(OutPutArray)* or *expression(OutPutArray)*

### Parameters

The required *OutPutArray* argument is an identifier specifying a dynamic `String` array variable.

### Return value

_None_

---

## Remarks

**Note**: *Before dump data, is recommended to make a `ImportFromCSV` method call.*

The *OutPutArray* parameter must be declared as dynamic `String` array. If user forget to do this, an error will occur.

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/docs/api/methods/importfromcsv.html).

---

## Behavior

The `DumpToArray` method make a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

The dumped data will be erased from memory, in other words, the current instance doesn't hold the CSV read data any more. In the same way, the `DumpToArray` method returns an empty `String` array for subsequent calls not preceded by `ImportFromCSV` method call.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/docs/api/methods/)