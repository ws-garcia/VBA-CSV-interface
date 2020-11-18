---
title: CommentsToken
parent: Properties
grand_parent: API
nav_order: 1
---

# CommentsToken
{: .d-inline-block }

New
{: .label .label-purple }

Gets or sets the char used for identify comments lines on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`CommentsToken`|
|Let|*expression*.`CommentsToken` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: Token:<br>*Type*: `String`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`String`|
|Let|_None_|

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>A line starting with the `CommentsToken` char is expected to be automatic ignored by the parser. By default, the char "#" is used for indicate commented lines, but this property can be set to whatever character. If the `CommentsToken` has a length greater than 1, only the first char of it is used.
{: .text-grey-dk-300 .bg-grey-lt-000 }

>⚠️**Caution**
>{: .text-grey-lt-000 .bg-green-000 }
>This property is only used when the `QuotingMode` property is set to `QuotationMode.Critical`.
{: .text-grey-dk-300 .bg-yellow-000 }

See also
: [EscapeToken Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapetoken.html), [QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [FieldsDelimiter Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)
