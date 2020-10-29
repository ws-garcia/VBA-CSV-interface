# ![VBA-CSV interface](/docs/assets/img/CSVinterface.png)
[![version](https://img.shields.io/static/v1?label=version&message=latest&color=brightgreen&style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/releases/latest/)
[![version](https://img.shields.io/static/v1?label=licence&message=GPL&color=informational&style=plastic)](https://www.gnu.org/licenses/gpl-3.0.html)

## Introductory words

VBA CSV interface is a class module developed to accomplish the data exchange task between VBA arrays and CSV files at high speed. Projects from [@sdkn104](https://github.com/sdkn104/VBA-CSV) and [@Senipah](https://github.com/Senipah/VBA-Better-Array), both on Github, were used for comparative performance purposes.

The parser is compatible with those CSV files compliant with the RFC-4180 standard, but add some useful features like: 
* In-line comments (with a user-defined character).
* Skip blanks lines and empty ones.
* User-defined escape character (option not available in _Power Query for Excel 2019_ and with some inconsistences when use the _From Text(Legacy)_ wizard)[[1]](#1).

<a id="1">[1]</a> 
Power Query, and its legacy counterpart, was not able to handle fields’ embedded line breaks when the CSV's "Text qualifier" is a Single Quote or the Apostrophe char.

## Getting started

If you don't know how to get started with VBA-CSV Interface class, visit the [documentation repo](https://ws-garcia.github.io/VBA-CSV-interface/).

## Benchmark

The benchmark results for VBA-CSV Interface are available at [this site](https://ws-garcia.github.io/VBA-CSV-interface/home/getting_started.html#benchmark).

## Licence

Copyright (C) 2020  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

