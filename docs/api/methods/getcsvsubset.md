---
title: GetCSVsubset
parent: Methods
grand_parent: API
nav_order: 10
---

# GetCSVsubset
{: .d-inline-block }

New
{: .label .label-purple }

Returns a set of records matching the criteria applied to a desired field from a CSV file.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`GetCSVsubset`*(filePath, filters, keyIndex, \[configObj:= Nothing\])*

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
<td style="text-align: left;"><em>filePath</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable representing the full path to the target CSV file.</td>
</tr>
<tr>
<td style="text-align: left;"><em>filters</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Variant</code> Type variable representing an array containing all the criteria to be applied to the desired field.</td>
</tr>
<tr>
<td style="text-align: left;"><em>keyIndex</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Long</code> Type variable representing the index of the field to apply the criteria.</td>
</tr>
<tr>
<td style="text-align: left;"><em>configObj</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>parserConfig</code> object variable holding the parser configuration.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `ECPArrayList` object

---

## Behavior

The `GetCSVsubset` method will retrieve all records where the field at position *keyIndex* meets all given criteria. If *configObj* is not given, the internal configuration of the current instance will be used.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>If the *filters* parameter is not an array, an error will occur.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
