---
title: QuotationMode
parent: Enumerations
grand_parent: API
nav_order: 2
---

# QuotationMode Enum
{: .fs-9 }

Provides a list of constants to configure the CSV parsing/writing operation behavior.
{: .fs-6 .fw-300 }

---

## Parts

|**_Constant_**|**_Member name_**|
|:----------|:----------|
|0|*Critical*|
|1|*All*|

---

## Syntax

*variable* = `QuotationMode`.*Constant*

---

:::note
The `QuotationMode.Critical` value, default one, is used to indicates the CSV file must use escape char only in fields having special char. The `QuotationMode.All` value most be used for those CSV files in wich all its fields will be escaped with the escape char given with the `EscapeChar` property.
:::

See also
: [EscapeChar Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html).

[Back to Enumerations overview](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/)