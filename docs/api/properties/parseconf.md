---
title: parseConfig
parent: Properties
grand_parent: API
nav_order: 12
---

# parseConfig
{: .fs-6 }

Holds the parser configuration for the current instance.
{: .fs-4 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`parseConfig`|
|Let|*expression*.`parseConfig` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: configuration<br>*Type*: `parserConfig`/`Object`<br>*Modifiers*: `ByRef`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`parserConfig`/`Object`|
|Let|_None_|

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
<td style="text-align: left; color:blue;"><em>bufferSize</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the buffer’s size, in MB, for the text stream operations using the <code>CSVTextStream.cls</code>. By default, this property is set to 0.5 MB for CSV/TSV file streams. When parsing a file with very long lines, the code attempts to adjust this value to avoid unexpected behavior.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>commentsToken</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the character used as commented line indicator. By default, the char <code>"#"</code> is used for indicate commented lines, but this property can be set to whatever single character. A line starting with the <code>commentsToken</code> char is ignored by the parser if the <code>skipCommentLines</code> is set to <code>True</code>. If the <code>commentsToken</code> has a length greater than 1, only the first char of it is used as indicator.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>CopyConfig</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns a <code>parserConfig</code> object with a copy of the current configuration.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>DefineTypingTemplate</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">This method is used to create the Dynamic Typing Template (DTT) through a <code>ParamArray</code>. User must specify a data conversion using the <code>TypeConversion</code> enumeration. Each DTT element must be linked to a field index.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>DefineTypingTemplateLinks</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">This method is used to link the Dynamic Typing Template (DTT), through a <code>ParamArray</code>, to specific fields. User must specify a column index (<code>Long</code>) for each element defined in the DTT.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>delimitersGuessing</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Sets the behavior of the parser regardless specified delimiters. By default, this property is set to <code>False</code>. If the value is set to <code>True</code> the parser will try to guess delimiters before start the import operation.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>delimitersToGuess</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the delimiters used on the <code>delimitersGuessing</code> operation. By default, the parser uses a <code>String</code> array with the chars comma (<code>,</code>), semicolon (<code>;</code>), Tab (<code>vbTab</code>),  pipe (<code>|</code>) and colon (<code>:</code>).</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dialect</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets a <code>CSVdialect</code> object with attributes to help the parser handle delimiters, quotes and escape modes in the requested CSV file. Refer to the documentation for the <code>CSVdialect</code> class for more detailed information on.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dTTemplateDefined</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;"> Gets the Dynamic Typing Template (DTT) status. The property returns <code>False</code> when the DTT was not defined.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dTTemplateLinksDefined</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;"> Gets the Dynamic Typing Template (DTT) Links status. The property returns <code>False</code> when the DTT links were not defined.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dTypingLinks</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the Dynamic Typing Template (DTT) Links through a <code>Variant</code> data type array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dTypingTemplate</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the Dynamic Typing Template (DTT) through a <code>Variant</code> data type array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>dynamicTyping</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the Dynamic Typing behavior. By default, this property is set to <code>False</code>. If the value is set to <code>True</code> the parser will use the DTT to type/convert the fields linked to the template.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>headers</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the header status of the target CSV/TSV file. By default, this property is set to <code>True</code>. When <code>False</code> the parser interpretates the target file hasn't header record.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>headersOmission</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">If <code>True</code>, the parser will omit the header record.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>multiEndOfLineCSV</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the behavior of the parser when reading streams from text files. By default, this property is set to <code>False</code>. If the value is set to <code>True</code>, all line break characters in the loaded stream will be converted to <code>vbLf</code>. This option will affect performance, but may be useful when faced with CSV files with <code>vbCrLf</code>, <code>vbCr</code> and <code>vbLf</code> mixed in as line endings.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>path</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">The full path, including the file name and its extension, to the target CSV/TSV.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>skipCommentLines</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the behavior of the parser when facing comments lines. By default, this property is set to <code>True</code>. If the value is set to <code>False</code>, the comment lines will be parsed as a normal record.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>skipEmptyLines</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the behavior of the parser when facing empty lines. By default, this property is set to <code>True</code>. If the value is set to <code>False</code>, the empty lines will be parsed as a normal record.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>startingRecord</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">This property must be used in combination with the <code>endingRecord</code> property for import a certain range of records from a CSV/TSV file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>utf8EncodedFile</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the behavior of the text stream reader when returning data read from a CSV file. By default, this property is set to <code>False</code>. If the value is set to <code>True</code>, the data obtained from the file will be interpreted as UTF-8 encoded and operated on before returning the buffer string. This property internationalizes the parser, making it capable of dealing with files in almost any foreign language: Chinese, Russian, Danish, Greek...</td>
</tr>
</tbody>
</table>

See also
: [CSVdialect class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvdialect.html), [CSVTextStream class](https://ws-garcia.github.io/VBA-CSV-interface/api/csvtextstream.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)