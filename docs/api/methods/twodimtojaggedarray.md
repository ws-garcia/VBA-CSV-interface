---
title: TwoDimToJaggedArray
parent: Methods
grand_parent: API
nav_order: 13
---

# TwoDimToJaggedArray
{: .fs-9 }

Deconstructs a 2D string array and puts its content into a jagged array.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`TwoDimToJaggedArray`*(TwoDimArray, JaggedArray)*

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
<td style="text-align: left;"><em>TwoDimArray</em></td>
<td style="text-align: left;">Required. Identifier specifying a dynamic <code>String</code> Type array variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>JaggedArray</em></td>
<td style="text-align: left;">Required. Identifier specifying a dynamic <code>Variant</code> Type array variable.</td>
</tr>
</tbody>
</table>

### Return value

_None_

---

## Behavior

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The *TwoDimArray* parameter must be declared as `String` type array, and will be successively deconstructed and erased by the `TwoDimToJaggedArray` method passing its content to the *JaggedArray* parameter.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)