---
title: ErrNumber
parent: Properties
grand_parent: API
---

# ErrNumber
{: .fs-9 }

Gets the number for the last occurred error over the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax
*expression*.`ErrNumber`

---

### Parameters

_None_

### Returns

*Type*: `Long`

---

## Remarks

Use the `ErrNumber` property to check if the last requested operation succeed.

See also: 
 [ErrDescription property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/errors/errdescription.html),
 [ErrSource property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/errors/errsource.html).
 
[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)

<table>
<thead>
<tr>
<th style="text-align: left;">Part</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>stringcheck</em></td>
<td style="text-align: left;">Required. <a href="../../glossary/vbe-glossary#string-expression" data-linktype="relative-path">String expression</a> being searched.</td>
</tr>
<tr>
<td style="text-align: left;"><em>stringmatch</em></td>
<td style="text-align: left;">Required. String expression being searched for.</td>
</tr>
<tr>
<td style="text-align: left;"><em>start</em></td>
<td style="text-align: left;">Optional. <a href="../../glossary/vbe-glossary#numeric-expression" data-linktype="relative-path">Numeric expression</a> that sets the starting position for each search. If omitted, -1 is used, which means that the search begins at the last character position. If <em>start</em> contains <a href="../../glossary/vbe-glossary#null" data-linktype="relative-path">Null</a>, an error occurs.</td>
</tr>
<tr>
<td style="text-align: left;"><em>compare</em></td>
<td style="text-align: left;">Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, a binary comparison is performed. See the Settings section for values.</td>
</tr>
</tbody>
</table>