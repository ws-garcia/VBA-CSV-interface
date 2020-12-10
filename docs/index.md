---
layout: default
title: Home
has_children: true
nav_order: 1
description: "VBA CSV interface is a class module that allows users exchange data between VBA arrays and CSV/TSV files at high speed."
---

# Introductory things
{: .fs-9 }

VBA CSV interface is a class module developed to accomplish the data exchange task between VBA arrays and CSV/TSV files at high speed. The class module doesn't use Excel Worksheets, neither any external reference such as MS Scripting Runtime.
{: .fs-6 .fw-300 }

## Advantages
* Writes and reads files at high speed.
* Supports those CSV's that follows the RFC-4180 specs.
* Supports Tab Separated Values (TSV) files. Gracefully handles line-breaks inside TSV fields enclosed in quotes.
* Allows individual access to imported fields and records in the VBA array style.
* Auto exclude any quote mark when data is imported.
* Allows an user-defined escape token (option not available in Power Query for Excel 2019 and with some inconsistences when user launch the From Text(Legacy) wizard). 
* Supports One-dimensional arrays, Two-dimensional arrays and jagged arrays.
* User has the option to import only certain range of records from given file.
* Supports in-line comments (with a user-defined character). See Papa Parse project.
* Supports blanks lines and empty ones.
