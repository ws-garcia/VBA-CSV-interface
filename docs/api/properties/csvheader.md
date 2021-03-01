---
title: csvHeader
parent: Properties
grand_parent: API
nav_order: 2
---

# csvHeader
{: .d-inline-block }

New
{: .label .label-purple }

Gets the imported CSV/TSV headers.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`csvHeader` 

### Parameters

_None_

### Returns

*Type*: `ECPArrayList`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The header record is read from a CSV/TSV file at importation time, if there is no data imported the headers are set to `Nothing`.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [parseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html).

---

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)