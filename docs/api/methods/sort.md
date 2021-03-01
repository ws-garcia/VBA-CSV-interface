---
title: Sort
parent: Methods
grand_parent: API
nav_order: 14
---

# Sort
{: .d-inline-block }

New
{: .label .label-purple }

Sorts the imported CSV/TSV data.

---

## Syntax

*expression*.`Sort`*(\[fromIndex:= -1\], \[toIndex:= -1\], \[SortColumn:= -1\], \[Descending:= False\])*

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
<td style="text-align: left;"><em>fromIndex</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>toIndex</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SortColumn</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Long</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Descending</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVinterface`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before sort data, is required to make a call to the `ImportFromCSV` or `ImportFromCSVstring` method.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

When the *fromIndex* parameter is omitted the data sorting start on the second record when the `ParseConfig.headers` is set to `True` and `ParseConfig.headersOmission` is set to `False`. Omitting the *toIndex* parameter takes the sorting to the last available record. Also, if the *SortColumn* parameter is omitted the data will sorted over the first column. Set the *Descending* parameter to `True` to sort the data in descending order.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
