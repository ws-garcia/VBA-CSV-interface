---
title: DumpToJaggedArray
parent: Methods
grand_parent: API
nav_order: 2
---

# DumpToJaggedArray
{: .d-inline-block }

New
{: .label .label-purple }

Dumps the data from the current instance to a jagged array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToJaggedArray`*(OutPutArray)*

### Parameters
The required *OutPutArray* argument is an identifier specifying a dynamic `Variant` type array variable.

### Return value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before dump data, is recommended to make a `ImportFromCSV` or `ImportFromCSVstring` method call. The *OutPutArray* parameter must be declared as dynamic `Variant` type array. If user forget to do this, an error can occur.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

The `DumpToJaggedArray` method make a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

The dumped data will be successively erased from memory, in other words, the current instance will doesn't hold the read CSV data any more. In the same way, the `DumpToJaggedArray` method doesn’t perform any modifications to the `String` type array for subsequent calls not preceded by one `ImportFromCSV` or `ImportFromCSVstring` method call.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `OutPutArray` will holds a set of `String` type arrays. To access to an individual element user must use something like **_expression(i)(j)_**, where **_i_** denotes an index in the main array and **_j_** denotes an index in the child array.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)