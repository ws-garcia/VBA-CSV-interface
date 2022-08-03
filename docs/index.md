---
layout: default
title: Home
has_children: true
nav_order: 1
description: "VBA CSV interface is a set of class modules that allows users exchange data between VBA arrays and CSV/TSV files."
---

# Introductory things
{: .fs-6 }

The most powerful and comprehensive CSV/[TSV](https://www.iana.org/assignments/media-types/text/tab-separated-values)/[DSV](https://www.linuxtopia.org/online_books/programming_books/art_of_unix_programming/ch05s02.html) data management library for VBA, providing parsing/writing capabilities compliant with RFC-4180 specifications and a complete set of tools for manipulating records and fields: [dedupe](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/dedupe.html), [sort](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/sort.html) and [filter](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/filter.html) records; [rearrange](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/rearrangefields.html), [shift](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/shiftfield.html), [merge](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/mergefields.html) and [split](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/splitfield.html) fields. Is your data spread over two or more CSV files? Don't worry, here you will find [Left, Right and Inner](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/tjoin.html) joins, and much more!
{: .fs-4 .fw-300 }

## Advantages
* __RFC-4180 specs compliant__.
* __Stable__. Fully Test Driven Developed (TDD) library, ([69/69 test passed](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/testing/tests/results/)), that includes 650+ line of code for testing. See [VBA test library by Tim Hall](https://github.com/ws-garcia/vba-test).
* __Fast__. Writes and reads files at the highest speed.
* __Memory-friendly__. Files are processed using a custom stream technique, only 0.5MB are in memory at a time.
* __Robust__. The library is not just a simple parser and writer, it is also a CSV data editor/manager.
* __[UTF-8](https://www.unicode.org/faq/utf_bom.html#UTF8) encoding support__. Do you have a CSV file, perhaps in chinese or some other foreign cyrillic language, downloaded from the Internet? This library is made to help you deal with it! You will be able to read and write UTF-8 encoded files in an easy way. 
* __Easy to use__. A few lines of code can do the work!
* __Automatic delimiter sniffer__. Don't worry if you forgot the file configuration. The interface has a solid strategy for sniff delimiters!
* __Highly Configurable__. User can configure the parser to work with a wide range of CSV files.
* __CSV data subsetting__. Split CSV data into a set of files with related data.
* __Like SQL queries on CSV files__. Use complex patterns to mimic SQL queries and filter data by criteria (=, <>, >=, <=, & (AND), \|(OR)).
* __Flexible__. Import only certain range of records from the given file, import fields (columns) by indexes or names, read records in sequential mode. 
* __Dynamic Typing support__. Turn CSV data field to a desired VBA data type.
* __Multi-level data sorting__. Sort CSV imported data over multiple columns using the hyper-fast(100k records per second) [Yaroslavskiy Dual-Pivot Quicksort](https://web.archive.org/web/20151002230717/http://iaroslavski.narod.ru/quicksort/DualPivotQuicksort.pdf) like Java and also other methods like: IntroSort, HeapSort and Merge sort.
* __Microsoft Access compatible__. The library has a version for those who feel in comfort working through DAO databases, [download from here](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip). 
