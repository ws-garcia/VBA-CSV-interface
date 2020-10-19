---
title: FieldsDelimiter
parent: Properties
grand_parent: API
nav_order: 8
---

# FieldsDelimiter
{: .fs-9 }

Indicates the char that will be used for delimit fields in the target CSV file.
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

---

## Remarks
The current behavior for the CSV interface forces `FieldsDelimiter` property to be distinct than a `Space` or`vbTab` char, set the property to comma or semicolon are the most logical options.

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)