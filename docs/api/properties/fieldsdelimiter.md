---
title: FieldsDelimiter
parent: Properties
grand_parent: API
nav_order: 9
---

# FieldsDelimiter
{: .fs-9 }

Indicates the char that will be used for delimit fields in the target CSV/TSV file.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`FieldsDelimiter`|
|Let|*expression*.`FieldsDelimiter` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: Delimiter:<br>*Type*: `String`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`String`|
|Let|_None_|

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>The current behavior for the CSV interface forces `FieldsDelimiter` property to be distinct than a `Space` char, set the property to comma, semicolon or `vbTab`, are the allowable options.
{: .text-grey-dk-300 .bg-yellow-000 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)