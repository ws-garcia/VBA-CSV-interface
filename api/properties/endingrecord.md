---
title: EndingRecord
parent: Properties
grand_parent: API
nav_order: 2
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

---

## Remarks
The `EndingRecord` property must be used in combination with the `StartingRecord` property for import a certain range of records from a desired CSV file.

See also
: [StartingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html).

---

## Behavior
The default value for the `EndingRecord` property is one(1) and force the class to import all the records from de CSV file starting at `StartingRecord` property. If the `EndingRecord` property is set to a value less than given on the `StartingRecord` property, only one record will be imported.

Setting the `EndingRecord` property to a value greater than the available records in the CSV file has the same effect than setting it to one(1).

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)