---
layout: default
title: Home
has_children: true
nav_order: 1
description: "VBA CSV interface is a class module that allows users exchange data between VBA arrays and CSV files at high speed."
---

# Introductory things
{: .fs-9 }

VBA CSV interface simplify the work with Comma Separated Value (CSV) files, allowing you to exchange information between an VBA array and an external CSV file without using Excel Worksheets, neither any external reference such as MS Scripting Runtime.
{: .fs-6 .fw-300 }

[Download now](https://github.com/ws-garcia/VBA-CSV-interface/releases/tag/v1.0.1){: .btn .btn-primary .fs-5 .mb-4 .mb-md-0 .mr-2 } [View it on GitHub](https://github.com/ws-garcia/VBA-CSV-interface){: .btn .fs-5 .mb-4 .mb-md-0 }

---

The CSV files are special kind of tabulated plain text data widely used in data exchange. There is no globally accepted standard format for that kind of files, however, out there are well formed standards such as [RFC4180](https://www.ietf.org/rfc/rfc4180.txt) proposed by The Internet Society.
Although many solutions has been developed for work with CSV files into VBA, including projects from [@sdkn104](https://github.com/sdkn104/VBA-CSV) and [@Senipah](https://github.com/Senipah/VBA-Better-Array) on GitHub, the performance philosophy conduce me to the development of a VBA class module that allows users exchange data between VBA arrays and CSV files at superior speed for the VBA programing language.

## Advantages
* Fully compliant with RFC4180 CSV standard.
* Exported data is 100% Excel spreadsheet compatible.
* Writes and reads files at high speed.
* Minimal Memory overload.
* User have the option to import only certain range of records from given CSV file.
* Auto exclude any quote mark when data is imported.
* Simple code logic that allows you easy modify and enhance it!

## Philosophy
The VBA CSVinterface class module is designed for gain advantage from the well-structured CSV files, this means, there isn't automatic syntax check, given the user decide how the class will works. This can be seen as a weakness, but the class get a speed-up on writing and reading procedures at time the user controls how the file is interpreted, keeping in mind that, in fact, VBA is a language with slow code execution speed. Under this idealization the developed solution complies with the RFC4180 standard for user specified CSV document format.

Keep in mind that class intentionally ignores the rule #7 of the RFC4180 standard [If double-quotes are used to enclose fields, then a double-quote appearing inside a field must be escaped by preceding it with another double quote], letting the users the chance to define the apostrophe as escape char.
{: .fs-4 .fw-300 }