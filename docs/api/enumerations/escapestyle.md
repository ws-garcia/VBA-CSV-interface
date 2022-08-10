---
title: EscapeStyle
parent: Enumerations
grand_parent: API
nav_order: 3
---

# EscapeStyle Enum
{: .fs-9 }

Provides a list of constants to configure the escape mechanism used when parsing/writing a CSV file.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|0|*rfc4180*|
|1|*unix*|

---

## Syntax

*variable* = `EscapeStyle`.*MemberName*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `EscapeStyle` enumeration is used in the `parseConfig.dialect.escapeMode` property to "escape" some CSV/TSV fields with embedded special characters. The `unix` constant will tell the parser to escape in the unix style, preceding the backslash (`\`), in contrast, `rfc4180` will escape fields as specified in RFC-4180, preceding another escape character. Be aware that if a field is escaped using RFC-4180, the embedded "\" characters will be part of the output and that the entire field will be surrounded by quotation marks (`'`, `"` or `~`).
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconfig.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)