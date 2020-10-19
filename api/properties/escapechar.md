---
title: EscapeChar
parent: Properties
grand_parent: API
---

# Expected CSV Escape char

## Description
Dictates the char that will be used for escape those fields containing some of the CSV syntax special char.
{: .fs-4 .fw-300 }

---

## Parts
ReadWrite: **_Yes_**{: .fs-4 .fw-300 }

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.**EscapeChar**|
|Let|*expression*.**EscapeChar** = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|**_None_**|
|Let|*Name*: **_EscapeChr_**:<br>*Type*: `EscapeType`/`Long`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`EscapeType`/`Long`|
|Let|**_None_**|

---

## Remarks
Setting the `EscapeChar` property to `EscapeType.NullChar` is only recommended when the `QuotingMode` property is set to `QuotationMode.All`. This scenario comes to reality when user have to work with CSV files over which neither fields need to be escaped.

The above means if the target CSV have an unknown structure, the best alternative is set the `EscapeChar` property to `EscapeType.DoubleQuotes` and the `QuotingMode` property to `QuotationMode.Critical`. These are the defaults settings.

See also:

[EscapeType Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetype.html).

[QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).
{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)