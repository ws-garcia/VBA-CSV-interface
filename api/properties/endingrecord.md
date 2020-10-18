---
title: EndingRecord
parent: Properties
grand_parent: API
---

# Importation ending record

## Description
Determines the record over which the import process will ends.
{: .fs-4 .fw-300 }

---

## Parts
ReadWrite: **Yes**{: .fs-4 .fw-300 }

---

## Syntax

|Accesor|Syntax|
|:----------|:----------|
|Get|*expression*.**EndingRecord**|
|Let|*expression*.**EndingRecord** = value|

{: .fs-4 .fw-300 }

|Accesor|Parameters|
|:----------|:----------|
|Get|**None**|
|Let|<table><thead></thead><tbody><tr><td>Name</td><td>Type</td><td>Modifiers</td></tr><tr><td>RecNumber</td><td>`Long`</td><td>ByVal</td></tr></tbody></table>|

{: .fs-4 .fw-300 }

|Accesor|Returns Type|
|:----------|:----------|
|Get|`Long`|
|Let|**None**|

{: .fs-4 .fw-300 }

---

## Remarks
Use the `EndingRecord` property in combination with the `StartingRecord` property for import a certain range of records from a desired CSV file.
{: .fs-4 .fw-300 }

[StartingRecord property overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html)

---
## Behavior
* The default value for the `EndingRecord` property is one(1) and force the class to import all the records from de CSV file starting at `StartingRecord` property.
* If the `EndingRecord` property is set to a value less than given on the `StartingRecord` property, only one record will be imported.
* Setting the `EndingRecord` property to a value greater than the available records in the CSV file has the same effect than setting it to one(1).
{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)