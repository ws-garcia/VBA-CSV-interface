---
title: QuotingMode
parent: Properties
grand_parent: API
nav_order: 13
---

# QuotingMode
{: .fs-9 }

Configures the CSV/TSV parsing/writing operation behavior.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`QuotingMode`|
|Let|*expression*.`QuotingMode` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: Mode<br>*Type*: `QuotationMode`/`Long`<br>*Modifiers*: `ByRef`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`QuotationMode`/`Long`|
|Let|_None_|

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>Use the `QuotingMode` property to set the parser behavior. Some files do not require further processes after an easy to do string split.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)