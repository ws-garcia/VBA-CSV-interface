---
layout: default
title: Dependencies
parent: Home
nav_order: 3
---

# Dependencies
{: .fs-6 }

The CSV interface library is composed of the following class modules:

* `CSVArrayList`. Developed to emulate some ArrayList functionalities present in other languages and to optimize operations on imported data.
* [`CSVcallBack`](https://github.com/ws-garcia/VBA-Expressions). Developed to offer a filtering path using Custom Functions.
* `CSVdialect`. To share the file dialect (delimiters and escape behavior) in a compact way.
* [`CSVexpressions`](https://github.com/ws-garcia/VBA-Expressions). The core filtering methods module, provides users with advanced filtering capabilities.
* `CSVinterface`. The main module for dealing with CSV file operations. 
* `CSVparserConfig`. To easily share all parser options between methods.
* `CSVSniffer`. CSV dialect guessing helper.
* `CSVTextStream`. To work with text files through streams, saving RAM memory.
* [`CSVudFunctions`](https://github.com/ws-garcia/VBA-Expressions). A place where users can write custom functions for use at runtime.