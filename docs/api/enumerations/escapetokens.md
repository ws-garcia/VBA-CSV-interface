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
|1|*Apostrophe*|
|2|*DoubleQuotes*|
|3|*Tilde*|

---

## Syntax

*variable* = `EscapeTokens`.*Constant*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `EscapeTokens` enumeration is used to "escape" some CSV/TSV fields with embedded escape characters. The `parserConfig.unixEscapeMechanism` option will tell the parser to escape in the unix style, preceding the backslash (`\`), or in the classic way, preceding another escape character.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconfig.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)