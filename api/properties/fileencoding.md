---
title: FileEncoding
parent: Properties
grand_parent: API
---

# FileEncoding
{: .fs-9 }

Returns the charset used to encode the last opened CSV file.
{: .fs-6 .fw-300 }

---

## ReadWrite

_ReadOnly_

---

## Syntax

*expression*.`FileEncoding`

### Parameters

_None_

### Returns

*Type*: `String`

---

## Remarks

The `FileEncoding` property is set when CSV file is load on memory. The property value could be one of this: ANSI, UTF-8, Unicode, BigEndian and Unknown.

Since VBA works with Unicode charset, a check to the `FileEncoding` property can help user overcome some codification issues. For this purposes, out there are free tools like [Notepad++](https://npp-user-manual.org/docs/preferences/) with options to change a file codification with just a left mouse click.

[Back to Properties overview](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/)