---
title: rectangularResults
parent: Properties
grand_parent: API
nav_order: 11
---

# rectangularResults
{: .d-inline-block }

New
{: .label .label-purple }

Gets the result array regularity status on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`rectangularResults`

### Parameters

_None_

### Returns

*Type*: `Boolean`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>If the `rectangularResults` property is `True`, the internal `ECPArrayList` is not irregular. A `False` value indicates the presence of, at least, one vector with more fields than the header record.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [vectorsBound property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/vectorsbound.html)

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)