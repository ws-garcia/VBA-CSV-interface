---
title: GetDataFromCSV
parent: Methods
grand_parent: API
nav_order: 4
---

# GetDataFromCSV
{: .fs-9 }

Dumps a CSV file content to a string variable
{: .fs-6 .fw-300 }

---

## Syntax

*expression*.`GetDataFromCSV`*(csvPathAndFilename)*

### Parameters

The required *csvPathAndFilename* argument is an identifier specifying a `String` variable.

### Return value

*Type*: `String`

>:pencil: **NOTE:**
>
>The *csvPathAndFilename* parameter must be the full path to the target CSV file, this means, the parameter holds the folder path, the file name and the ".csv" extension.

## Behavior

The `GetDataFromCSV` method returns an empty `String` when errors occurs.

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)