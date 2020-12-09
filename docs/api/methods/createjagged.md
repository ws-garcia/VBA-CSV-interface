---
title: CreateJagged
parent: Methods
grand_parent: API
nav_order: 1
---

# CreateJagged
{: .fs-9 }

Creates an empty array of vectors, each of which having a fixed custom size.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`CreateJagged`*(ArrVar, ArraySize, VectorSize)*

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
<td style="text-align: left;"><em>ArrVar</em></td>
<td style="text-align: left;">Required. Identifier specifying a dynamic <code>Variant</code> Type array variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>ArraySize</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Must be greater or equal to zero.</td>
</tr>
<tr>
<td style="text-align: left;"><em>VectorSize</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable. Must be greater or equal to zero.</td>
</tr>
</tbody>
</table>

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>Setting the *ArraySize* or the *VectorSize* parameter to a value less than zero will generate a runtime error.
{: .text-grey-dk-300 .bg-yellow-000 }

### Return value

_None_

---

## Behavior

The *ArraySize* parameter is used by the `CreateJagged` method to resize the *ArrVar* array. In the same way, the *VectorSize* parameter is used for set the sizes of `String` Type vectors.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>To access to an individual element user must use something like **_expression(i)(j)_**, where **_i_** denotes an index in the main array and **_j_** denotes an index in the child array.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)