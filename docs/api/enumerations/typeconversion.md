---
title: TypeConversion
parent: Enumerations
grand_parent: API
nav_order: 3
---

# TypeConversion Enum
{: .d-inline-block }

New
{: .label .label-purple }

Provides a list of constants to configure the Dynamic Typing conversion behavior.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|-1|*None*|
|0|*ToLong*|
|1|*ToDouble*|
|2|*ToDate*|
|3|*ToBoolean*|

---

## Syntax

*variable* = `TypeConversion`.*Constant*

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `TypeConversion.None` value is used to indicates the specified CSV/TSV file field is a `String` data type and not need conversion to other type.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [ParseConfig Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconfig.html), [Import Example \[EXAMPLE2\]](https://ws-garcia.github.io/VBA-CSV-interface/examples/importation-examples.html#example2).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)