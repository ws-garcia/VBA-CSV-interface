---
title: EndLineChar
parent: Enumerations
grand_parent: API
nav_order: 2
---

# EndLineChar Enum
{: .fs-9 }

Provides a list of constants to configure the `WriteBlankLines` method of the `CSVTextStream.cls` module.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|0|*CRLF*|
|1|*CR*|
|2|*LF*|

---

## Syntax

*variable* = `EndLineChar`.*MemberName*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>By default, the `WriteBlankLines` use `EndLineChar.CRLF` value to appends lines to the target CSV/TSV file.
{: .text-grey-dk-300 .bg-grey-lt-000 }

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)