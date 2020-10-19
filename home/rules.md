---
layout: default
title: Rules
parent: Home
nav_order: 3
---

# RFC-4180 rules
{: .fs-9 }

The CSVinterface class is intended to be fully RFC-4180 CSV standard compliant, despite this some tweaks are added to make the interface more robust and purposes useful.

In the table bellow all the rules of [RFC-4180](https://www.ietf.org/rfc/rfc4180.txt) standard and its counterparts on the `CSVinterface.cls` are listed. This topic is highly recommended for CSV interface behavior knowledge.

|*RFC-4180 rule*|*Over CSV interface*|
|:--------------------------------------------------|:--------------------------------------------------|
|*Each record is located on a separate line, delimited<br> by a linebreak (CRLF).*|Accepts also CR or LF instead.|
|*The last record in the file may or may not have an<br> ending line break.*|In the same way. Includes a routine for avoid read<br> empty lines.|
|*There maybe an optional header line appearing as<br> the first line of the file with the same format<br> as normal record lines.  This header will contain<br> names corresponding to the fields in the file and<br> should contain the same number of fields as the<br> records in the rest of the file.*|In the same way. The presence or absence of the<br> header line should be indicated via the optional<br> "header".|
|*Within the header and each record, there may be<br> one or more fields, separated by commas.  Each<br> line should contain the same number of fields<br> throughout the file.  Spaces are considered part<br> of a field and should not be ignored.  The last<br> field in the record must not be followed by a<br> comma.*|In the same way. The spaces betwen the fields<br> separator char and a single filed is ignored<br> only if that filed need to be escaped.|
|*Each field may or may not be enclosed in double<br> quotes (however some programs, such as Microsoft<br> Excel, do not use double quotes at all).  If<br> fields are not enclosed with double quotes, then<br> double quotes may not appear inside the fields*|In the same way. The class accepts also the<br> apostrophe char for indicate fields needing to<br> be escaped. It's important to notice that a<br> single CSV record may have fields enclosed and<br> not enclosed by the escape char.|
|*Fields containing line breaks (CRLF), double<br> quotes, and commas should be enclosed in double<br> quotes*|In the same way. Also accepts fields enclosed by<br> the apostrophe char.|
|*If double-quotes are used to enclose fields, then<br> a double-quote appearing inside a field must be<br> escaped by preceding it with another double quote.*|Ignored rule. The class accepts the apostrophe<br> as escape char, and follow the standard claims<br> may cause conflict with some abbreviate US<br> slangs (i.e.: "**_isn't_**").|