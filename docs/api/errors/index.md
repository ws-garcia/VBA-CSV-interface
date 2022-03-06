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
<td style="text-align: left;">[CSV file Export]: The passed argument isn't an array or a CSVArrayList object.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212491</em></td>
<td style="text-align: left;">[CSV file subset]: The specified CSV file is empty. No subset can be processed.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212489</em></td>
<td style="text-align: left;">[CSV Field Insert]: Cannot insert a field in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212488</em></td>
<td style="text-align: left;">The specified index is out of bounds. Please check and try again.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212487</em></td>
<td style="text-align: left;">[CSV Field Remove]: Cannot remove the field in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<td style="text-align: left;"><em>-2147212486</em></td>
<td style="text-align: left;">[CSV Record Insert]: Cannot insert the record in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212485</em></td>
<td style="text-align: left;">[CSV Records Remove]: Cannot remove the records in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212484</em></td>
<td style="text-align: left;">[CSV Fields Rearrange]: Cannot rearrange the fields in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212483</em></td>
<td style="text-align: left;">[CSV Fields Rearrange]: The order specified for the fields is incomplete. Please enter all field indexes before continuing.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212482</em></td>
<td style="text-align: left;">[CSV Fields Merge]: Cannot merge the fields in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212481</em></td>
<td style="text-align: left;">[CSV Field Split]: Cannot split the field in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212480</em></td>
<td style="text-align: left;">[CSV Field Shift: Cannot shift the field in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
<tr>
<td style="text-align: left;"><em>-2147212479</em></td>
<td style="text-align: left;">[CSV Field Shift: Cannot shift the record in the current instance. This is because there is no imported data or the records do not have the same number of fields.</td>
<td style="text-align: left;">CSVinterface Class.</td>
</tr>
</tbody>
</table>