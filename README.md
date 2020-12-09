# ![VBA-CSV interface](/docs/assets/img/CSVinterface.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/releases/latest)

## Introductory words

VBA CSV interface is a class module developed to accomplish the data exchange task between VBA arrays and CSV files at high speed. Projects from [@sdkn104](https://github.com/sdkn104/VBA-CSV) and [@Senipah](https://github.com/Senipah/VBA-Better-Array), both on Github, were used for comparative performance purposes.

## Advantages
* Writes and reads files at high speed.
* Supports those CSV's that follows the RFC-4180 specs.
* Supports [Tab Separated Values (TSV)](https://www.iana.org/assignments/media-types/text/tab-separated-values) files. Gracefully handles line-breaks inside TSV fields enclosed in quotes.
* Allows individual access to imported fields and records in the VBA array style.
* Auto exclude any quote mark when data is imported.
* Allows an user-defined escape token (option not available in _Power Query for Excel 2019_ and with some inconsistences when user launch the _From Text(Legacy)_ wizard)[[1]](#1).
* Supports One-dimensional arrays, Two-dimensional arrays and [jagged arrays](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/arrays/jagged-arrays).
* User has the option to import only certain range of records from given file.
* Supports in-line comments (with a user-defined character). See [Papa Parse](https://www.papaparse.com/) project.
* Supports blanks lines and empty ones.

<a id="1">[1]</a> 
Power Query, and its legacy counterpart, was not able to handle fields’ embedded line breaks when the CSV's "Text qualifier" is a Single Quote or the Apostrophe char.

## Getting started

If you don't know how to get started with VBA-CSV Interface class, visit the [documentation repo](https://ws-garcia.github.io/VBA-CSV-interface/).

## Contributing

In order to contribute whit in this project, please see the [guidance for contributing](https://ws-garcia.github.io/VBA-CSV-interface/contributing.html).

## Benchmark

The benchmark results for VBA-CSV Interface are available at [this site](https://ws-garcia.github.io/VBA-CSV-interface/home/getting_started.html#benchmark).

##Limitations

Visit [this site](https://ws-garcia.github.io/VBA-CSV-interface/limitations/csv_file_size.html) in order to known the around CSV file size considerations.

## Licence

Copyright (C) 2020  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

