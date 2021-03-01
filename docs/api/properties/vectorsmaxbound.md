---
title: vectorsMaxBound
parent: Properties
grand_parent: API
nav_order: 12
---

# vectorsMaxBound
{: .d-inline-block }

New
{: .label .label-purple }

Gets the maximum bound of the vectors in the result array on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`vectorsMaxBound`

### Parameters

_None_

### Returns

*Type*: `Long`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `vectorsMaxBound` property returns the max number of fields, concerning to one record, into the result array. This property is useful to reserve memory for copy data from the internal `ECPArrayList` to a 2D array.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [vectorsBound property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/vectorsbound.html), [rectangularResults property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/rectangularresults.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)