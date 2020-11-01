---
title: EndingRecord
parent: Properties
grand_parent: API
nav_order: 3
---

# EndingRecord
{: .fs-9 }

Determines the record over which the import process will ends.
{: .fs-6 .fw-300 }

---

## ReadWrite

_Yes_

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.`EndingRecord`|
|Let|*expression*.`EndingRecord` = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|_None_|
|Let|*Name*: RecNumber:<br>*Type*: `Long`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`Long`|
|Let|_None_|

>📝**Note**
>{: .text-grey-lt-000 .bg-green-000 }
>The `EndingRecord` property must be used in combination with the `StartingRecord` property for import a certain range of records from a CSV file.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html).

---

## Behavior

The `EndingRecord` property can be set to an value greater or equal than zero (0), trying to set it to a negative value will setting the property to its default value.

The default value for the `EndingRecord` property is one(1) and force the class to import all successive records from de CSV file starting at the record specified with the `StartingRecord` property. Setting the `EndingRecord` property to zero (0) and the `StartingRecord` property to one (1) , will import the first record from the CSV file only. Whatever other `EndingRecord` property value less than the given in the `StartingRecord` property forces import one record.

Setting the `EndingRecord` property to a value greater than the available records in the CSV file has the same effect than setting it to one(1).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)
