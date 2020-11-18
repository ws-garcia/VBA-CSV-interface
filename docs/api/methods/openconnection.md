---
title: OpenConnection
parent: Methods
grand_parent: API
nav_order: 11
---

# OpenConnection
{: .fs-9 }

Loads a CSV/TSV file on memory for data Input/Output operations.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`OpenConnection`*(csvPathAndFilename, {DeleExistingFile:= `False`})*

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
<td style="text-align: left;"><em>csvPathAndFilename</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable holding the file's path.</td>
</tr>
<tr>
<td style="text-align: left;"><em>DeleExistingFile</em></td>
<td style="text-align: left;">Optional. Identifier specifying a <code>Boolean</code> Type variable.</td>
</tr>
</tbody>
</table>

### Return value

_None_

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The `OpenConnection` method don't rejects any kind of file extension, user need to ensure the target file has a name ending in `.csv`, `.tsv` or `.txt`.
{: .text-grey-dk-300 .bg-yellow-000 }

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `OpenConnection` method is the preamble to the `ImportFromCSV` and `ExportToCSV` methods, this means each call to the citated methods must be preceded by a `OpenConnection` method call.
>After call the `OpenConnection` method is possible to check if the instance is bind to the target file, for which is only needed to read the current instance `Connected` property.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [Connected property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/connected.html), [ImportFromCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html), [ExportToCSV method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/exporttocsv.html).

---

## Behavior

When the given path exists the file will be created on that path, otherwise an error occur. For on path existing file, the `OpenConnection` method will delete the file when the *DeleExistingFile* parameter is set to `True`. If that is not the case, a new file will be created.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)