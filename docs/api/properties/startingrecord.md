---
title: StartingRecord
parent: Properties
grand_parent: API
nav_order: 15
---

# StartingRecord
{: .fs-9 }

Determines the record over which the import process will starts.
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
>Use the `StartingRecord` property in combination with the `EndingRecord` property for import a certain range of records from CSV file.
{: .text-grey-dk-300 .bg-grey-lt-000 }

See also
: [EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html).

## Behavior

The `StartingRecord` property can be set to an value greater or equal than one (1), trying to set it to a value less than the cited one will setting the property to its default value. 
 
The default value for the `StartingRecord` property is one(1) and force the class to start the importation over the first available record in the CSV file. If the `StartingRecord` property is set to a value greater than the available records in the CSV file, neither record will be imported.

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)
