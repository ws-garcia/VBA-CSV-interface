 ---
title: CommentLineIndicator
parent: Properties
grand_parent: API
nav_order: 1
---

# CommentLineIndicator
{: .fs-9 }

Gets or sets the char used for identify comments lines on the current instance.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`CommentLineIndicator`|
|Let|*expression*.`CommentLineIndicator` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: RecNumber:<br>*Type*: `String`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`String`|
|Let|_None_|

---

>:pencil: **NOTE:**
>
>A line starting with the `CommentLineIndicator` char is expected to be automatic ignored by the parser. By default, the char "#" is used for indicate commented lines, but this property can be set to whatever character. If the `CommentLineIndicator` has a length greater than 1, only the first char of it is used.

>:warning: **CAUTION**
>
>This option is only available when the `QuotingMode` property is set to `QuotationMode.Critical`.

See also
: [EscapeChar Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/escapechar.html), [QuotationMode Enumeration](https://ws-garcia.github.io/VBA-CSV-interface/api/enumerations/quotationmode.html), [FieldsDelimiter Property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/fieldsdelimiter.html).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)