---
title: DumpToAccessTable
parent: Methods
grand_parent: API
nav_order: 3
---

# DumpToAccessTable
{: .d-inline-block }

New
{: .label .label-purple }

Dumps the data from the current instance to a Microsoft Access Database (.accdb).
{: .fs-6 .fw-300 }

---

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>This method is only available in the [Access version of the CSVinterface.cls](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip) module.
{: .text-grey-dk-300 .bg-yellow-000 }

## Syntax

*expression*.`DumpToAccessTable`*(dBase, tableName, \[fieldsToIndexing\])*

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
<td style="text-align: left;"><em>dBase</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>DAO.Database</code> object variable representing the output database.</td>
</tr>
<tr>
<td style="text-align: left;"><em>tableName</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable representing the output database table name.</td>
</tr>
<tr>
<td style="text-align: left;"><em>fieldsToIndexing</em></td>
<td style="text-align: left;">Optional. Identifiers specifying a <code>Variant</code> Type variables representing the fields of the table in where indexing is required.</td>
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

If a table named *tableName* already exist in the database, the operation is aborted. User can specify the fields for create indexes by name or by absolute position through the *fieldsToIndexing* parameter. When the *fieldsToIndexing* parameter is not set, the data is dumped into the database table without indexing more fields than the record position.

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>All the data is dumped as "Short Text". If the CSV file has some of the special chars listed in [this article](https://docs.microsoft.com/en-us/office/troubleshoot/access/error-using-special-characters) an error can occur.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
