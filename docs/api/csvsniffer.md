---
title: CSVSniffer
parent: API
nav_order: 7
---

# CSVSniffer
{: .fs-6 }

Class module developed as an attempt to sniff/guess CSV dialects without user intervention. In some preliminary tests, the sniffer was 100% accurate, but there is always the risk of facing ambiguous cases that can only be solved with human intervention. This class is inspired by the [work of scientist Till Roman Döhmen](https://homepages.cwi.nl/~boncz/msc/2016-Doehmen.pdf), with some improvements to disambiguate the most complicated cases.
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
<td style="text-align: left; color:blue;"><em>DetectDataType</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Attempts to detect the data type of a CSV field. The method can detect numeric, alphanumeric, currency, date and time, email, file system paths, IP v4, percentages, urls, structured data from programming languages (bytearray, frozenset, JS arrays). The method will return <code>1</code> when it can recognize the data type present in the specified field and <code>0</code> when the field contains an unknown data type.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>TableScore</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Calculates a score for the CSV data based on the congruence of the detected data type and the uniformity of the fields contained in each record. The score is in the range <code>0 < x <= 100</code>. The higher the score obtained, the higher the probability that the dialect used is the correct one for the data in the analyzed CSV file. The user can pass as <code>ArrayList</code> parameter the imported data or the Items stored through the <code>Add2</code> method.</td>
</tr>
</tbody>
</table>

[Back to API overview](https://ws-garcia.github.io/VBA-CSV-interface/api/)