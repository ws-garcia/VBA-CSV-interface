---
title: ExportToCSV
parent: Methods
grand_parent: API
nav_order: 9
---

# ExportToCSV
{: .fs-9 }

Exports an array's content to a CSV/TSV file.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`ExportToCSV`*(csvArray, \[pconfig:= `Nothing`\], \[PassControlToOS:= `True`\], \[enableDelimiterGuessing:= `True`\])*

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
<td style="text-align: left;"><em>csvArray</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Variant</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>pconfig</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>ParserConfig</code> object variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>PassControlToOS</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
<tr>
<td style="text-align: left;"><em>enableDelimiterGuessing</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
</tbody>
</table>

### Returns value

_None_

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The *csvArray* parameter can be an `ECPArrayList` or an array (one-dimensional, two-dimensional or jagged) variable, passing another type of variable will cause an error. 
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html).

---

## Behavior

If the `pconfig` parameter is omited, the parser will use the `ParseConfig` property as configuration object. If the file specified in the `.path` configuration property all ready exist and have some content on it the parser will try to guess delimiters and the data will be append to the file.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>If there is an unescaped field, the `ExportToCSV` method will append the specified escape character to the end of the field. This is because if the data is written with this problem, the parser will not be able to import it on future occasions.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
