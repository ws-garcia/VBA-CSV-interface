---
title: DumpToSheet
parent: Methods
grand_parent: API
nav_order: 2
---

# DumpToSheet
{: .fs-9 }

Dumps the data from the current instance to an Excel WorkSheet.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToSheet`*({WBookName}, {SheetName}, {RngName})*

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
<td style="text-align: left;"><em>WBookName</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> variable representing the output Workbook name.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SheetName</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> variable representing the output Worksheet name.</td>
</tr>
<tr>
<td style="text-align: left;"><em>RngName</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> variable representing the name of the output top left-most range.</td>
</tr>
</tbody>
</table>

### Return value

_None_

>:pencil: **NOTE:**
>
>Before dump data, is recommended to make a `ImportFromCSV` or `ImportFromCSVstring` method call.

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

When the *WBookName* parameter is omitted the data is dumped into the Workbook that holds the CSV interface's *VBAProject*. Omitting the *SheetName* parameter adds a new Worksheet to the desired Workbook. Also, if the *RngName* parameter is omitted the data will dumped starting on the "A1" named cell in the desired Worksheet.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)