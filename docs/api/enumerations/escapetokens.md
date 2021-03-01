---
title: EscapeTokens
parent: Enumerations
grand_parent: API
nav_order: 2
---

# EscapeTokens Enum
{: .fs-9 }

Provides a list of constants to configure the char used as escape character.
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

*variable* = `EscapeTokens`.*Constant*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `EscapeTokens.NullChar` value is used to indicates the CSV/TSV file does not use any escape char in its whole length. This value induces the program to write the file assuming the `parserConfig.fieldsDelimiter` property is enough for the export operation.
>
>In the case the `parserConfig.fieldsDelimiter` property is not enough for successfully done the export operation, the `EscapeTokens.DoubleQuotes` value would be used for parse/write an CSV/TSV having fields to be escaped with double quote and the `EscapeTokens.Apostrophe` values for parse/write a file having fields to be escaped with the apostrophe.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconfig.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)