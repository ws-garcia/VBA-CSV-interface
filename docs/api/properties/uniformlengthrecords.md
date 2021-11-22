---
title: uniformLengthRecords
parent: Properties
grand_parent: API
nav_order: 13
---

# uniformLengthRecords
{: .d-inline-block }

New
{: .label .label-purple }

Gets the result array regularity status on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`uniformLengthRecords`

### Parameters

_None_

### Returns

*Type*: `Boolean`

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>If the `uniformLengthRecords` property is `True`, the internal `CSVArrayList` is not irregular. A `False` value indicates the presence of at least one record with more or fewer fields than the header record.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [fieldsBound property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsbound.html)

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)