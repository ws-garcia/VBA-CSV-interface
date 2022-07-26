---
title: CSVTextStream
parent: API
nav_order: 8
---

# CSVTextStream
{: .fs-6 }

Easy-to-use class module developed to enable I/O operations over "big" text files, at high speed, from VBA. The module hasn’t reference to any external API library and has the ability to read and write UTF-8 encoded files.
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
<td style="text-align: left; color:blue;"><em>atEndOfStream</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the overall state of the pointer on the text stream. Returns <code>True</code> if the file pointer is at the end of a file, and <code>False</code> otherwise.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>bufferLength</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the number of string characters in the buffer.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>bufferSize</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the buffer size, in MB, for text stream operations. Allows the user to specify how much data is read at a time. By default, the bufferSize property is set to 0.5 MB. For files containing very long lines, the size is modified to be enough to contain at least one line.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>bufferString</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the text data stored in the buffer. If one or both of the <code>unifiedLFOutput</code> and <code>utf8EncodedFile</code> properties are set to <code>True</code>, the string is operated on before returning data.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>CloseStream</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Closes the current text file stream. After close the current stream, user will lose the connection to CSV file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>endStreamOnLineBreak</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Allows to end buffer just after the first, from right to left, line break character.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>isOpenStream</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the stream status over the current CSV file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>lineBreak</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the last line break character read.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>linebreakMatchingBehavior</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the behavior of line break matching in subsequent text stream operations. Allows the user to specify how line breaks are searched for to end the current buffer in <code>vbCrLf</code>, <code>vbCr</code> or <code>vbLf</code> as specified in the <code>endStreamOnLineBreak</code> property. By default, the property property is set to <code>EndLineMatchingBehavior.Bidirectional</code>, this option ensures the handling of files with long lines that cannot be contained in a string of specified size as in the <code>bufferSize</code> property. Setting the <code>linebreakMatchingBehavior</code> property to <code>EndLineMatchingBehavior.OnlyBackwardSense</code> may cause unexpected behavior when the stream is requested to end at a line break and the current text stream contains a portion of a long line that cannot be stored in the specified buffer size.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>OpenStream</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Opens a stream over a CSV file. Before stream over a text file, user must open a stream pointing to that file. If the text file doesn’t exist, it will be created and then a stream is opened.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>pointerPosition</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the overall pointer position over the current text file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ReadText</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reads a number of characters from the stream file and saves the result to the <code>bufferString</code> of current instance. Each call to this method will read a set of characters until the buffer size is reached. If the <code>EndStreamOnLineBreak</code> property is set to <code>True</code>, the stream will be cut off at the first occurrence of a line break (<code>CRLF</code>, <code>CR</code> or <code>LF</code>) in the reverse left (right-to-left) direction or some extra data will be appended to it until a line break character is encountered in the forward direction. The <code>ReadText</code> method will continue to read data until the pointer exceeds the length of the current text file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RestartPointer</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Moves the pointer to the initial position of the CSV file and clears the buffer. The user must open a stream before attempting to restart the reader.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>SeekPointer</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Moves the pointer, over the target file, to the specified position. The next I/O operation will start in the position specified with the <code>SeekPointer</code> method.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>streamLength</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets the current opened file’s size, in Bytes.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>unifiedLFOutput</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Determines whether the buffer string is returned using only the <code>vbLf</code> character as end of lines.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>UTF8Decode</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Decodes an UTF-8 string. This method makes it possible to read and write CSV files in foreign languages, share mathematical symbols through CSV files and much more.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>utf8EncodedFile</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Determines whether the buffer string is decoded from UTF-8 encoding.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>WriteBlankLines</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Inserts a specified number of blank lines into the current opened CSV file.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>WriteText</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Writes the given string to the current opened CSV file. User must open a stream before try to write data to file.</td>
</tr>
</tbody>
</table>

[Back to API overview](https://ws-garcia.github.io/VBA-CSV-interface/api/)