---
title: DumpToJaggedArray
parent: Methods
grand_parent: API
nav_order: 5
---

# DumpToJaggedArray
{: .fs-9 }

Dumps the data from the current instance to a jagged array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToJaggedArray`*(OutPutArray)*

### Parameters

The required *OutPutArray* argument is an identifier specifying a dynamic `Variant` type array variable.

### Returns value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before dump data, is required to make a call to the `ImportFromCSV` or `ImportFromCSVstring` method. The *OutPutArray* parameter must be declared as dynamic `Variant` type array. If user forget to do this, an error can occur.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

The `DumpToJaggedArray` method makes a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The *OutPutArray* argument will contain a set of `Variant` type arrays. To access to an individual element user must use something like **_expression(i)(j)_**, where **_i_** denotes an index in the main array and **_j_** denotes an index in the child array.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
