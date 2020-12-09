---
title: Item
parent: Properties
grand_parent: API
nav_order: 14
---

# Item
{: .d-inline-block }

New
{: .label .label-purple }

Gets a field, or an array with an entire record, from the result array on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`Item`*(Index1, \[Index2\])*

### Parameters

<table>
<thead>
<tr>
<th style="text-align: left;">Part</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>Index1</em></td>
<td style="text-align: left;">Required. Identifier specifying a numeric Type value representing the position, in the first dimension of the result array, over which the requested data will be retrieved.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Index2</em></td>
<td style="text-align: left;">Optional. Identifier specifying a numeric Type value representing the position, over a vector from the result array, on which the requested data will be retrieved.</td>
</tr>
</tbody>
</table>

### Returns

*Type*: `Variant`/`String`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `Item` property can be used for retrieve information from the class internals when only certain range of the parsed CSV data is needed. Don’t use this property when working with big files (up to 350 MB), there is potential risk of run out of memory if you are using VBA x64 and 8 GB of RAM.
>
>If user only provide the *Index1* as argument, an entire record will be returned; if user provide more than one argument, the *Index2* will be used to return a field.
{: .text-grey-dk-300 .bg-grey-lt-000 }

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The user must check, through the `RectangularResults` and the `IrregularVectors` properties, if the read CSV has records with varying number of fields. This step can prevent potential "subscript out of range" error.
{: .text-grey-dk-300 .bg-yellow-000 }

See also
: [VectorsBound property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/vectorsbound.html), [RectangularResults property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/rectangularresults.html), [IrregularVectors property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/irregularvectors.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)