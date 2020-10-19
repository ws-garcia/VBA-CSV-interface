---
title: EscapeType
parent: Enumerations
grand_parent: API
nav_order: 1
---

# EscapeType Enum
{: .fs-9 }

Provides a list of constants for use to configure the char used as escape one.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|0|*NullChar*|
|1|*Apostrophe*|
|2|*DoubleQuotes*|

---

## Syntax

*variable* = `EscapeType`.*Constant*

---

## Remarks

The `EscapeType.NullChar` value is used with the`QuotationMode.All` setting to indicates the CSV file does not use any escape char in its whole length. This values combination conduces the CSV file to be parse/write assuming the `FieldsDelimiter` property is enough for the import/export operations.

In the case the `FieldsDelimiter` property is not enough for successfully done the import/export operations, the `QuotationMode.DoubleQuotes` value would be used for parse/write an CSV having fields to be escaped with double quote and the `QuotationMode.Apostrophe` values for parse/write an CSV having fields to be escaped with the apostrophe. 

See also:
 [EscapeChar Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [FieldsDelimiter Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)