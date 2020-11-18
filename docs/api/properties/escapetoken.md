---
title: EscapeToken
parent: Properties
grand_parent: API
nav_order: 7
---

# EscapeToken
{: .fs-9 }

Dictates the char that will be used for escape those fields containing some of the CSV/TSV syntax special char.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`EscapeToken`|
|Let|*expression*.`EscapeToken` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: Token:<br>*Type*: `EscapeTokens`/`Long`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`EscapeTokens`/`Long`|
|Let|_None_|

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Setting the `EscapeToken` property to `EscapeTokens.NullChar` is only recommended when the `QuotingMode` property is set to `QuotationMode.All`. This scenario comes to reality when user have to work with files over which neither fields need to be escaped.
>
>The above means if the target file have an unknown structure, the best alternative is set the `EscapeToken` property to `EscapeTokens.DoubleQuotes` and the `QuotingMode` property to `QuotationMode.Critical`. These are the defaults settings.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [EscapeTokens Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/escapetokens.html), [QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)