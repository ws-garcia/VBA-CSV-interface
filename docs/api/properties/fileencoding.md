---
title: FileEncoding
parent: Properties
grand_parent: API
nav_order: 9
---

# FileEncoding
{: .fs-9 }

Returns the charset used to encode the last opened CSV file.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`FileEncoding`

### Parameters

_None_

### Returns

*Type*: `String`

---

## Remarks

The `FileEncoding` property is set when CSV file is load on memory and can return the following values:

<table>
<thead>
<tr>
<th style="text-align: left;">Value</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>ANSI</em></td>
<td style="text-align: left;">When the last opened CSV file is encoded in ANSI<br> format. Each char is represented using 8 bits.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Unicode</em></td>
<td style="text-align: left;">The last opened CSV file is encoded in Unicode<br> format. Combinations of basic chars are allowed<br> over the CSV file and its takes 2 bytes per<br> character.</td>
</tr>
<tr>
<td style="text-align: left;"><em>UTF-8</em></td>
<td style="text-align: left;">The CSV file is encoded in UTF-8 format.</td>
</tr>
<tr>
<td style="text-align: left;"><em>BigEndian</em></td>
<td style="text-align: left;">The charset encoding follows the Most Significative<br> Byte (MSB) logic.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Unknown</em></td>
<td style="text-align: left;">The charset encoding can’t be determinate.</td>
</tr>
</tbody>
</table>

Since VBA works with Unicode charset, a check to the `FileEncoding` property can help user overcome some codification issues. For this purposes, out there are free tools like [Notepad++](https://npp-user-manual.org/docs/preferences/) with options to change a file codification with just a left mouse click.

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/docs/api/properties/)