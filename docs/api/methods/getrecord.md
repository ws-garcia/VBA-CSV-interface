---
title: GetRecord
parent: Methods
grand_parent: API
nav_order: 11
---

# GetRecord
{: .d-inline-block }

New
{: .label .label-purple }

Reads a new record from the CSV sequentially forward.
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`GetRecord`

### Parameters

_None_

### Returns value

*Type*: `ECPArrayList`

See also
: [CloseSeqReader Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/closeseqreader.html), [OpenSeqReader Method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/openseqreader.html).

---

## Behavior

The `GetRecord` method returns an `ECPArrayList` object containing the data of a CSV record. If an error occurs or the end of file (EOF) is reached, the method returns `Nothing`.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)