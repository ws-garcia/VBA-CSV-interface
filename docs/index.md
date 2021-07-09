---
layout: default
title: Home
has_children: true
nav_order: 1
description: "VBA CSV interface is a set of class modules that allows users exchange data between VBA arrays and CSV/TSV files."
---

# Introductory things
{: .fs-9 }

VBA CSV interface is the most complete, and open source, CSV/TSV VBA parser library nowadays. The library is RFC-4180 compliant and enables users to manipulate CSV content at the highest speed. All the modules were developed to accomplish the data exchange task with the greatest performance and to grant an easy use.
{: .fs-6 .fw-300 }

## Advantages
* __RFC-4180 specs compliant__.
* __Stable__. Fully Test Driven Developed (TDD) library, ([63/63 test passed](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/testing/tests/results/)), that includes 650+ line of code for testing. See [VBA test library by Tim Hall](https://github.com/ws-garcia/vba-test).
* __Fast__. Writes and reads files at the highest speed.
* __Memory-friendly__. CSV/[TSV](https://www.iana.org/assignments/media-types/text/tab-separated-values) files are processed using a custom stream technique, only 0.5MB are in memory at a time.
* __Robust__. Parser and writer accept [Unix-style quotes escape sequences](https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml#notes). 
* __Easy to use__. A few lines of code can do the work!
* __Automatic delimiter guesser__. Don't worry if you forgot the file configuration. The interface has a solid strategy for guessing delimiters!
* __Highly Configurable__. User can configure the parser to work with a wide range of CSV files.
* __CSV data subsetting__. Split CSV data into a set of files with related data.
* __Like SQL queries on CSV files__. Add your own logic to mimic SQL queries and filter data by criteria (=, <>, >=, <=, AND, OR).
* __Flexible__. Import only certain range of records from the given file, import fields (columns) by indexes or names, read records in sequential mode. 
* __Dynamic Typing support__. Turn CSV data field to a desired VBA data type.
* __Data sorting__. Sort CSV imported data using the hyper-fast(100k records per second) [Yaroslavskiy Dual-Pivot Quicksort](https://web.archive.org/web/20151002230717/http://iaroslavski.narod.ru/quicksort/DualPivotQuicksort.pdf) like Java.
* __Microsoft Access compatible__. The library has a version for those who feel in comfort working through DAO databases, [download from here](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip).