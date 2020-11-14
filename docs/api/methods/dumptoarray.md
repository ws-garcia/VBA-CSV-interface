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

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before dump data, is recommended to make a `ImportFromCSV` or `ImportFromCSVstring` method call. The *OutPutArray* parameter must be declared as dynamic `String` array. If user forget to do this, an error can occur.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

The `DumpToArray` method make a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

The dumped data will be successively erased from memory, in other words, the current instance will doesn't hold the CSV read data any more. In the same way, the `DumpToArray` method doesn’t perform any modifications to the `String` type array for subsequent calls not preceded by one `ImportFromCSV` or `ImportFromCSVstring` method call.
[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)