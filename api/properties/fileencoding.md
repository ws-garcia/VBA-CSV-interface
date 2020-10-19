---
title: FileEncoding
parent: Properties
grand_parent: API
---

# Encoding charset

## Description
Returns the charset used to encode the last opened CSV file.
{: .fs-4 .fw-300 }

## Parts
ReadWrite: **_ReadOnly_**{: .fs-4 .fw-300 }

## Syntax
*expression*.**FileEncoding**{: .fs-4 .fw-300 }

### Parameters

**_None_**{: .fs-4 .fw-300 }

### Returns

*Type*: `String`{: .fs-4 .fw-300 }

## Remarks
The `FileEncoding` property is set when CSV file is load on memory. The property value could be one of this: ANSI, UTF-8, Unicode, BigEndian and Unknown.

Since VBA works with Unicode charset, a check to the `FileEncoding` property can help user overcome some codification issues. For this purposes, out there are free tools like [Notepad++](https://notepad-plus-plus.org) that can change a file codification with just a left mouse click.
{: .fs-4 .fw-300 }

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)