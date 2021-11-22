---
layout: default
title: Dependencies
parent: Home
nav_order: 3
---

# Dependencies
{: .fs-9 }

The CSV interface library is composed of the following class modules:

* `CSVArrayList`. Developed to emulate some ArrayList functionalities present in other languages and to optimize operations on imported data.
* `CSVdialect`. To share the file dialect (delimiters and escape behavior) in a compact way.
* `CSVinterface`. The main module for dealing with CSV file operations. 
* `CSVparserConfig`. To easily share all parser options between methods.
* `CSVSniffer`. CSV dialect guessing helper.
* `CSVTextStream`. To work with text files through streams, saving RAM memory.