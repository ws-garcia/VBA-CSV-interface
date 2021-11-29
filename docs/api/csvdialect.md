---
title: CSVdialect
parent: API
nav_order: 6
---

# CSVdialect
{: .d-inline-block }

New
{: .label .label-purple }

Class module developed to share CSV dialects, or group of specific and related configuration, which instructs the parser on how to interpret the character set read from a CSV file. This container travels through the parsing and sniffer methods.
{: .fs-4 .fw-300 }

---

## Members

<table>
<thead>
<tr>
<th style="text-align: left;">Item</th>
<th style="text-align: left;">Type</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left; color:blue;"><em>escapeMode</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the escape behavior. The default value of this property is <code>EscapeStyle.rfc4180</code> and instructs the parser to escape fields according to RFC-4180 specs, if <code>EscapeStyle.unix</code> is used it will instruct the parser to escape special characters ALSO in Unix style by preceding those characters with a backslash <code>\</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>fieldsDelimiter</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the character to be used to delimit the fields in the target CSV/TSV file. The default value of this property is comma <code>,</code> but the user can specify any character as field delimiter. However, it is advisable to use commonly used characters such as colon <code>:</code>, semicolon <code>;</code>, pipe <code>|</code> and tab <code>vbTab</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>quoteToken</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the char that will be used for quote those fields containing some CSV/TSV syntax special char. The user must use the <code>QuoteTokens</code> enumeration to define this property. The user can choose between double quotes <code>"</code>, single quote <code>'</code> and tilde <code>~</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>recordsDelimiter</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the char that will be used for delimit records in the target CSV/TSV file. The default value is <code>vbCrLf</code>, but user can choose one of <code>vbCr</code> and <code>vbLf</code>.</td>
</tr>
</tbody>
</table>

[Back to API overview](https://ws-garcia.github.io/VBA-CSV-interface/api/)