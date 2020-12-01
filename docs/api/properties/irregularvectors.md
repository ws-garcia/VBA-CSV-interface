---
title: IrregularVectors
parent: Properties
grand_parent: API
nav_order: 13
---

# IrregularVectors
{: .fs-9 }

Gets a collection of arrays with INFO for irregular vectors on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`IrregularVectors`

### Parameters

_None_

### Returns

*Type*: `Collection`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `IrregularVectors` property returns a `Collection` of arrays with information for those vectors having a number of fields greater than those defined in the `VectorsBound` property for the current instance.
>
>Here "irregular" is used to denote the vectors that have more fields than the header record. Each array in the returned collection is zero based and contains a pair of values: index (position of the irregular vector in the result data array) and the number of fields at that index.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [VectorsBound property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/vectorsbound.html), [RectangularResults property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/rectangularresults.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)