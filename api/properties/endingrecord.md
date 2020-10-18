---
title: EndingRecord
parent: Properties
grand_parent: API
---

# Importation ending record

## Description
Determines the record over which the import process will ends.

## Parts
ReadWrite: **Yes**{: .fs-4 .fw-300 }

## Syntax
*expression*.**EndingRecord**{: .fs-4 .fw-300 }

### Parameters

**None**{: .fs-4 .fw-300 }

### Returns

Type: `Long`{: .fs-4 .fw-300 }

## Remarks
Use the `EndingRecord` property in combination with the `StartingRecord` property for import a certain range of records from a desired CSV file.{: .fs-4 .fw-300 }
[StartingRecord property overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/startingrecord.html)

## Behavior
* The default value for the `EndingRecord` property is one(1) and force the class to import all the records from de CSV file starting at `StartingRecord` property.{: .fs-4 .fw-300 }
* If the `EndingRecord` property is set to a value less than given on the `StartingRecord` property, only one record will be imported.{: .fs-4 .fw-300 }
* Setting the `EndingRecord` property to a value greater than the available records in the CSV file has the same effect than setting it to one(1).{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)