# ![VBA-CSV interface](/docs/assets/img/CSVinterface.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/releases/latest)

## Introductory words

VBA CSV interface is the most complete, and open source, CSV/TSV VBA parser library nowadays. The library is RFC-4180 compliant and enables users to manipulate CSV content at the highest speed. All the modules were developed to accomplish the data exchange task with the greatest performance and to grant an easy use.

## Advantages
* __Stable__. Fully Test Driven Developed (TDD) library, ([49/49 test passed](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/testing/tests/results/)), that includes 500+ line of code for testing. See [VBA test library by Tim Hall](https://github.com/ws-garcia/vba-test).
* __Fast__. Writes and reads files at the highest speed.
* __Memory-friendly__. CSV/[TSV](https://www.iana.org/assignments/media-types/text/tab-separated-values) files are processed using a custom stream technique, only 0.5MB are in memory at a time.
* __Robust__. Parser and writer accept [Unix-style quotes escape sequences](https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml#useful).
* __Easy to use__. A few lines of code can do the work!
* __Highly Configurable__. User can configure the parser to work with a wide range of CSV files.
* __CSV data subsetting__. Split CSV data into a set of files with related data.
* __Like SQL queries on CSV files__. Add your own logic to mimic SQL queries and filter data by criteria (=, <>, >=, <=, AND, OR).
* __Automatic delimiter guesser__. Don't worry if you forgot the file configuration!
* __Flexible__. Import only certain range of records from the given file, import fields (columns) by indexes or names, read records in sequential mode. 
* __Dynamic Typing support__. Turn CSV data field to a desired VBA data type.
* __Data sorting__. Sort CSV imported data using the hyper-fast(100k records per second) [Yaroslavskiy Dual-Pivot Quicksort](https://web.archive.org/web/20151002230717/http://iaroslavski.narod.ru/quicksort/DualPivotQuicksort.pdf) like Java.
* __Microsoft Access compatible__. The library has a version for those who feel in comfort working through DAO databases, [download from here](https://github.com/ws-garcia/VBA-CSV-interface/raw/master/src/Access_version.zip).
* __RFC-4180 specs compliant__.

## Getting started

If you don't know how to get started with VBA-CSV Interface class, visit the [documentation repo](https://ws-garcia.github.io/VBA-CSV-interface/). For code hints, basic and more in-depth use of the library, see [examples](https://ws-garcia.github.io/VBA-CSV-interface/examples/).

Visit the [frequently asked questions section](https://ws-garcia.github.io/VBA-CSV-interface/home/FAQ.html) for the most common questions.

## Contributing

In order to contribute within this project, please see the [guidance for contributing](https://ws-garcia.github.io/VBA-CSV-interface/contributing.html).

## Benchmark

The benchmark results for VBA-CSV Interface are available at [this site](https://ws-garcia.github.io/VBA-CSV-interface/home/getting_started.html#benchmark).

## Dependencies

The library depends on the [ECPTextStream class]( https://github.com/ws-garcia/ECPTextStream) in order to work with text files. In the same way, the class uses two external class modules: one for configuration sharing and the other for data storage. All dependencies are written in pure VBA.

## Limitations

Visit [this site](https://ws-garcia.github.io/VBA-CSV-interface/limitations/csv_file_size.html) in order to known around CSV file size considerations.

## Licence

Copyright (C) 2021  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

