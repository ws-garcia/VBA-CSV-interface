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

*expression*.`DumpToArray` *(OutPutArray)*
 or
 *expression (OutPutArray)*

### Parameters

The required *OutPutArray* argument is an identifier specifying a dynamic `String` array variable.

### Return value

_None_

---

## Remarks

The *OutPutArray* parameter must be declared as dynamic string array. If user forget to do this, an error will occur.

---

## Behavior

The `DumpToArray` method make a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)