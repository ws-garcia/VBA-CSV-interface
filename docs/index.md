---
layout: default
title: Home
has_children: true
nav_order: 1
description: "VBA CSV interface is a set of class modules that allows users exchange data between VBA arrays and CSV/TSV files."
---

# Introductory things
{: .fs-9 }

VBA CSV interface is the most complete, and open source, CSV/TSV VBA parser library nowadays. The library is RFC-4180 compliant and enables users to manipulate CSV content at the highest speed. All the modules were developed to accomplish the data exchange with the task with the greatest performance and to grant an easy use.
{: .fs-6 .fw-300 }

## Advantages
* Fast. Writes and reads files at the highest speed.
* Memory-friendly. CSV/[TSV](https://www.iana.org/assignments/media-types/text/tab-separated-values) files are processed using a custom stream technique, only 0.5MB are in memory at a time.
* Easy to use. A few lines of code can do the work!
* Highly Configurable. User can configure the parser to work with a wide range of CSV files.
* Automatic delimiter guesser. Don't worry if you forgot the file configuration!
* Flexible. Import only certain range of records from the given file, import fields (columns) by indexes or names.
* Dynamic Typing support. Turn CSV data field to a desired VBA data type.
* Data sorting. Sort CSV imported data using the hyper-fast [Yaroslavskiy Dual-Pivot Quicksort](https://web.archive.org/web/20151002230717/http://iaroslavski.narod.ru/quicksort/DualPivotQuicksort.pdf) like Java.
* RFC-4180 specs compliant.
* Auto skip blanks lines and empty ones.
* Supports in-line comments (with a user-defined character). See [Papa Parse](https://www.papaparse.com/) project. 
* Acces to the imported data in the VBA array style.
* Supports One-dimensional arrays, Two-dimensional arrays and [jagged arrays](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/arrays/jagged-arrays).
