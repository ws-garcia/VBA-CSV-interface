---
title: RecordsDelimiter
parent: Properties
grand_parent: API
---

# Expected CSV records delimiter char

## Description
Indicates the char that will be used for delimit records in the target CSV file.
{: .fs-4 .fw-300 }

---

## Parts
ReadWrite: **_Yes_**{: .fs-4 .fw-300 }

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.**RecordsDelimiter**|
|Let|*expression*.**RecordsDelimiter** = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|**_None_**|
|Let|*Name*: **_Delimiter_**:<br>*Type*: `String`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`String`|
|Let|**_None_**|

---

## Remarks
The `RecordsDelimiter` property can be set to`vbCr`, `vbCrLf` or `vbLf`. This options unlocks a limitation from RFC-4180 CSV standard .

{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)