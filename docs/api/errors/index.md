---
title: Errors
parent: API
has_children: false
nav_order: 4
---

# CSV interface errors documentation
{: .fs-9 }

This section describes the custom errors whose components are returned by the parser in the `errNumber`, `errSource` and `errDescription` properties. Please note that VBA-specific errors may occur during operations, the documentation of which is provided by Microsoft.

<table>
<thead>
<tr>
<th style="text-align: left;">Error number</th>
<th style="text-align: left;">Error description</th>
<th style="text-align: left;">Error source</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>-2147212498</em></td>
<td style="text-align: left;">Missing some escape char. Check the data and try again. [Review the record #?, field #? on the source CSV file/string].</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212497</em></td>
<td style="text-align: left;">The config object has an invalid Dynamic Typing Template (DTT). The number of Dynamic Typing Links (DTL) must be less or equal than the number of Dynamic Typing Targets Fields (DTTF) defined.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212496</em></td>
<td style="text-align: left;">The config object is not linked to a CSV file. Ensure set the path property to valid CSV before import data.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212494</em></td>
<td style="text-align: left;">The CSV file/String has no significant data. This can occur when the file/String has only empty or commented lines that can be omitted.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212493</em></td>
<td style="text-align: left;">The specified source CSV/String is empty. Please check and try again.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212492</em></td>
<td style="text-align: left;">[CSV file Export]: The passed argument isn't an array or a ECPArrayList object.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212491</em></td>
<td style="text-align: left;">[CSV file subset]: The specified CSV file is empty. No subset can be processed.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212490</em></td>
<td style="text-align: left;">[CSV file subset]: The given path name is an empty string or the specified CSV file does not exist in the supplied path.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
</tbody>
</table>