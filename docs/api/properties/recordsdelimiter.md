---
title: RecordsDelimiter
parent: Properties
grand_parent: API
nav_order: 14
---

# RecordsDelimiter
{: .fs-9 }

Indicates the char that will be used for delimit records in the target CSV/TSV file.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`RecordsDelimiter`|
|Let|*expression*.`RecordsDelimiter` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: Delimiter<br>*Type*: `String`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`String`|
|Let|_None_|

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `RecordsDelimiter` property can be set to`vbCr`, `vbCrLf` or `vbLf`. This options unlocks a limitation from the RFC-4180 specs.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)