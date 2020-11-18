---
layout: default
title: Rules
parent: Home
nav_order: 3
---

# RFC-4180 rules
{: .fs-9 }

The CSVinterface class is intended to be nearest as possible to the RFC-4180 specs for CSV files, despite this some tweaks are added to make the interface more robust and useful.

In the table bellow all the rules of [RFC-4180](https://www.ietf.org/rfc/rfc4180.txt) specs and its counterparts on the `CSVinterface.cls` are listed. This topic is highly recommended for CSVinterface behavior knowledge.

<table>
<thead>
<tr>
<th style="text-align: left;">RFC-4180 rule</th>
<th style="text-align: left;">Over CSV interface</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>Each record is located on a separate line, delimited<br> by a linebreak (CRLF).</em></td>
<td style="text-align: left;">Accepts also CR or LF instead.</td>
</tr>
<tr>
<td style="text-align: left;"><em>The last record in the file may or may not have an<br> ending line break.</em></td>
<td style="text-align: left;">In the same way. Includes a routine for avoid read<br> empty lines.</td>
</tr>
<tr>
<td style="text-align: left;"><em>There maybe an optional header line appearing as<br> the first line of the file with the same format<br> as normal record lines.  This header will contain<br> names corresponding to the fields in the file and<br> should contain the same number of fields as the<br> records in the rest of the file.</em></td>
<td style="text-align: left;">In the same way. The presence or absence of the<br> header line should be indicated via the optional<br> "HeadersOmission" parameter.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Within the header and each record, there may be<br> one or more fields, separated by commas.  Each<br> line should contain the same number of fields<br> throughout the file.  Spaces are considered part<br> of a field and should not be ignored.  The last<br> field in the record must not be followed by a<br> comma.</em></td>
<td style="text-align: left;">The class accepts CSV files with different numbers<br> of fields per record. The spaces betwen the<br> fields separator char and a single field is ignored<br> only if that field is enclosed in double quotes.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Each field may or may not be enclosed in double<br> quotes (however some programs, such as Microsoft<br> Excel, do not use double quotes at all).  If<br> fields are not enclosed with double quotes, then<br> double quotes may not appear inside the fields</em></td>
<td style="text-align: left;">In the same way. The class accepts also the<br> apostrophe char for indicate fields needing to<br> be escaped. It's important to notice that a<br> single CSV record may have fields enclosed and<br> not enclosed by the escape char.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Fields containing line breaks (CRLF), double<br> quotes, and commas should be enclosed in double<br> quotes</em></td>
<td style="text-align: left;">In the same way. Also accepts fields enclosed by<br> the apostrophe char.</td>
</tr>
<tr>
<td style="text-align: left;"><em>If double-quotes are used to enclose fields, then<br> a double-quote appearing inside a field must be<br> escaped by preceding it with another double quote.</em></td>
<td style="text-align: left;">Ignored rule. The class accepts the apostrophe<br> as escape char, and follow the specs claims<br> may cause conflict with some abbreviate US<br> slangs (e.g.: "<strong>isn't</strong>").</td>
</tr>
</tbody>
</table>
