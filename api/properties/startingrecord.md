---
title: StartingRecord
parent: Properties
grand_parent: API
---

# Importation starting record

## Description
Determines the record over which the import process will starts.
{: .fs-4 .fw-300 }

---

## Parts
ReadWrite: **_Yes_**{: .fs-4 .fw-300 }

---

## Syntax

|**_Accesor_**|**_Syntax_**|
|:----------|:----------|
|Get|*expression*.**EndingRecord**|
|Let|*expression*.**EndingRecord** = value|

|**_Accesor_**|**_Parameters_**|
|:----------|:----------|
|Get|**_None_**|
|Let|*Name*: **_RecNumber_**:<br>*Type*: `Long`<br>*Modifiers*: `ByVal`|

|**_Accesor_**|**_Returns Type_**|
|:----------|:----------|
|Get|`Long`|
|Let|**_None_**|

---

## Remarks
Use the `StartingRecord` property in combination with the `EndingRecord` property for import a certain range of records from a desired CSV file.
{: .fs-4 .fw-300 }

See also:

[EndingRecord property](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/endingrecord.html)

---

## Behavior
The default value for the `StartingRecord` property is one(1) and force the class to start the importation over the first available record in the CSV file. If the `StartingRecord` property is set to a value greater than the available records in the CSV file, neither record will be imported.
{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)