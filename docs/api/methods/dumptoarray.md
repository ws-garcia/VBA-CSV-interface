---
title: DumpToArray
parent: Methods
grand_parent: API
nav_order: 4
---

# DumpToArray
{: .fs-9 }

Dumps the data from the current instance to an array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToArray`*(OutPutArray)*

### Parameters

The required *OutPutArray* argument is an identifier specifying a dynamic `String` type array variable.

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

The `DumpToArray` method makes a copy of all the data stored in the current instance. The data is returned in the *OutPutArray* parameter for avoid additional data copies in the internals.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The data is always returned in a Two-dimensional array, even when the imported file only contain a field per record.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
