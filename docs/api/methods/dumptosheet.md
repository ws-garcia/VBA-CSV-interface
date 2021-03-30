---
title: DumpToSheet
parent: Methods
grand_parent: API
nav_order: 8
---

# DumpToSheet
{: .fs-9 }

Dumps data from a source, or from the current instance, to an Excel WorkSheet.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`DumpToSheet`*(\[WBookName\], \[SheetName\], \[RngName:= "A1"\], \[DataSource:= `Nothing`\])*

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
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> Type variable representing the output Workbook name.</td>
</tr>
<tr>
<td style="text-align: left;"><em>SheetName</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> Type variable representing the output Worksheet name.</td>
</tr>
<tr>
<td style="text-align: left;"><em>RngName</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>String</code> Type variable representing the name of the output top left-most range.</td>
</tr>
<tr>
<td style="text-align: left;"><em>DataSource</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>ECPArrayList</code> object variable representing the data to copy from.</td>
</tr>
</tbody>
</table>

### Returns value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Before dump data, is required to make a call to the `ImportFromCSV` or `ImportFromCSVstring` method.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ImportFromCSVstring method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html).

---

## Behavior

When the *WBookName* parameter is omitted the data is dumped into the Workbook that holds the CSV interface's *VBAProject*. Omitting the *SheetName* parameter adds a new Worksheet to the desired Workbook. Also, if the *RngName* parameter is omitted the data will dumped starting on the "A1" named cell in the desired Worksheet.

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>When the *DataSource* parameter is omitted the `DumpToSheet` method dumps all data stored in the current instance. If the user specified a data source, its data will be dumped.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
