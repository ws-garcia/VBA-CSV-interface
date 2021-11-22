# ![VBA-CSV interface](/docs/assets/img/CSVinterface.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-CSV-interface?style=plastic)](https://github.com/ws-garcia/VBA-CSV-interface/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/ws-garcia/VBA-CSV-interface/total.svg)](https://github.com/ws-garcia/VBA-CSV-interface/releases/)
[![Follow](https://img.shields.io/github/followers/ws-garcia.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/ws-garcia/VBA-CSV-interface/watchers)

## Introductory words

VBA CSV interface is the most complete, and open source, CSV/TSV VBA parser library nowadays. The library is RFC-4180 compliant and enables users to manipulate CSV content at the highest speed. All the modules were developed to accomplish the data exchange task with the greatest performance and to grant an easy use.

## Advantages
* __RFC-4180 specs compliant__.
* __Stable__. Fully Test Driven Developed (TDD) library, ([68/68 test passed](https://github.com/ws-garcia/VBA-CSV-interface/blob/master/testing/tests/results/)), that includes 650+ line of code for testing. See [VBA test library by Tim Hall](https://github.com/ws-garcia/vba-test).
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

## Getting started

If you don't know how to get started with VBA-CSV Interface class, visit the [documentation repo](https://ws-garcia.github.io/VBA-CSV-interface/). For code hints, basic and more in-depth use of the library, see [examples](https://ws-garcia.github.io/VBA-CSV-interface/examples/).

Visit the [frequently asked questions section](https://ws-garcia.github.io/VBA-CSV-interface/home/FAQ.html) for the most common questions.

### Using the Code

This section will attempt to analyze all the capabilities of the CSV interface

Import whole CSV file:

```
Sub CSVimport()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"                ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","         ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf     ' Rows delimiter
    End With
    With CSVint
        .ImportFromCSV .parseConfig             ' Import the CSV to internal object
    End With
End Sub
```

Now suppose from the file *"Sample.csv"* the user only requires to import a specific range of records. It is possible to write a code like the one shown below:

```
Sub CSVimportRecordsRange()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"                ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","         ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf     ' Rows delimiter
        .startingRecord = 10                   ' Start import on the tenth record
        .endingRecord = 20                     ' End of importation in the 20th record
    End With
    With CSVint
        .ImportFromCSV .parseConfig             ' Import the CSV to internal object
    End With
End Sub
```

If the user wants to sort the imported data, a code like the following can be written:

```
Sub CSVimportAndSort()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"               ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","                ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf            ' Rows delimiter
    End With
    With CSVint
        .ImportFromCSV .parseConfig           ' Import the CSV to internal object
        .Sort SortingKeys:=-1                 ' Sort imported data on first column is descending order
    End With
End Sub
```

CSV data are mainly treated as text strings, what if the user wants to do some calculations on the data obtained from a given file? In this situation, the user can change the behavior of the parser to work in dynamic typing mode. Here's an example:

```
Sub CSVimportAndTypeData()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"                       ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","                ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf            ' Rows delimiter
        .dynamicTyping = True                         ' Enable dynamic typing mode
        '@---------------------------------------------------------
        ' Configure dynamic typing
        .DefineTypingTemplate TypeConversion.ToDate, _
                                TypeConversion.ToLong, _
                                TypeConversion.ToDouble
        .DefineTypingTemplateLinks 6, _
                                    7, _
                                    10
        ' The dynamic typing mode will perform the following:
        '      * Over column 6 ---> String To Date data Type conversion
        '      * Over column 7 ---> String To Long data Type conversion
        '      * Over column 10 ---> String To Double data Type conversion
    End With
    With CSVint
        .ImportFromCSV .parseConfig             ' Import the CSV to internal object
    End With
End Sub
```

The quote character, used for escape fields, can be defined as one of them, according to an enumeration:

```
Sub SetQuoteChar()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig.dialect
        .quoteToken = QuoteTokens.DoubleQuotes  ' 2 = ["] (Default)
        '.quoteToken = QuoteTokens.Apostrophe   ' 1 = [']
        '.quoteToken = QuoteTokens.Tilde        ' 3 = [~]
    End With
End Sub
```

Once the data is imported and saved to the internal object, the user can access it in the same way as a standard VBA array. An example would be:

```
Sub LoopData(ByRef CSVint As CSVinterface)
    With CSVint
        Dim iCounter As Long
        Dim cRecord() As Variant              ' Records are stored as a one-dimensional array.
        Dim cField As Variant
        
        For iCounter = 0 To CSVint.count - 1
            cRecord() = .item(iCounter)       ' Retrieves a record
            cField = .item(iCounter, 1)       ' Retrieves the 2nd field of the current record
        Next
    End With
End Sub
```

In addition, the user can use the [`CSVArrayList`](https://ws-garcia.github.io/VBA-CSV-interface/api/csvarraylist.html) class to access the contents using code like this:

```
Sub LoopData2(ByRef CSVint As CSVinterface)
    With CSVint
        Dim iCounter As Long
        Dim cRecord() As Variant                    ' Records are stored as a one-dimensional array.
        Dim cField As Variant
        
        For iCounter = 0 To CSVint.count - 1
            cRecord() = .items.item(iCounter)       ' Retrieves a record
            cField = .items.item(iCounter)(1)       ' Retrieves the 2nd field of the current record
        Next
    End With
End Sub
```

However, it is sometimes disadvantageous to store data in containers other than VBA arrays. This becomes especially noticeable when it is required to write the information stored in Excel's own objects, such as spreadsheets, or VBA user forms, the case of list boxes, which allow to be filled in a single instruction using arrays. Then, the user can copy the information from the internal object using code like this:

```
Sub DumpData(ByRef CSVint As CSVinterface)
    Dim oArray() As Variant
    With CSVint
        .DumpToArray oArray            ' Dump the internal data into a two-dimensional array
        .DumpToJaggedArray oArray      ' Dump the internal data into a jagged array
        oArray = .items.items          ' Dump the internal data into a jagged array
        .DumpToSheet                   ' Dump the internal data into a new sheet
                                       ' using ThisWorkbook
        '@-------------------------------------------------------------------
        ' *NOTE: ONLY AVAILABLE FOR THE ACCESS VERSION OF THE CSV INTERFACE
        ' Dump the internal data into the Table1 in oAccessDB database.
        ' The method would create indexes in the 2nd and 3th fields.
        .DumpToAccessTable oAccessDB, _
                           "Table1", _
                            2, 3
    End With
End Sub
```

So far, in the examples addressed, the user has been allowed to choose between two actions:

<ol>
	<li>Import <em>ALL records</em> contained in a CSV file.</li>
	<li>Import a <em>recordset</em>, starting at record X and ending at record Y.</li>
</ol>

In both options, the user is forced to import all fields (columns) present in the file. Most CSV file parsers only offer the first option, but what if the user wants to save only the information that is relevant for a particular purpose, and what if user wants to store in memory only the records that meet a certain set of requirements?

An user may need to import 2 of X columns from a CSV file, in this case, the user can use something like:

```
Sub CSVimportDesiredColumns()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"              ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","       ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf   ' Rows delimiter
    End With
    With CSVint
        .ImportFromCSV .parseConfig, _
                        1, "Revenue"         ' Import 1st and "Revenue" fields ONLY
    End With
End Sub
```

So, OK, let's imagine now that an user wants to apply some logic before saving the data, in which case they can step through the records in the CSV file one by one, using the sequential reader, as shown in the following example:

```
Sub CSVsequentialImport()
    Dim CSVint As CSVinterface
    Dim csvRecord As CSVArrayList
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"             ' Full path to the file, including its extension.
        .dialect.fieldsDelimiter = ","      ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf  ' Rows delimiter
    End With
    With CSVint
        .OpenSeqReader .parseConfig, _
                        1, "Revenue"        ' Import the 1st and "Revenue" fields using
                                            ' seq. reader
        Do
            Set csvRecord = .GetRecord
            '//////////////////////////////////////////////
            'Implement your logic here
            '//////////////////////////////////////////////
        Loop While Not csvRecord Is Nothing   ' Loop until the end of the file is reached
    End With
End Sub
```

Is there a way to sequentially fetch a set of records at a time instead of a single record? Currently, there is no built-in method to do that with a single instruction, as in the examples above, but with a few extra lines of code and the tools provided by the library, it is possible to achieve that goal. This is illustrated in the following example where the CSV file is streamed:

```
Sub CSVimportChunks()
    Dim CSVint As CSVinterface
    Dim StreamReader As CSVTextStream
            
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .dialect.fieldsDelimiter = ","                      ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf                  ' Rows delimiter
    End With
    Set StreamReader = New CSVTextStream
    With StreamReader
        .endStreamOnLineBreak = True                        ' Instruct to find line breaks
        .OpenStream "C:\Sample.csv"                         ' Connect to CSV file
        Do
            .ReadText                                       ' Read a CSV chunk
            CSVint.ImportFromCSVString .bufferString, _
                                    CSVint.parseConfig, _
                                    1, "Revenue"            ' Import a set of records
            '//////////////////////////////////////
            'Implement your logic here
            '//////////////////////////////////////
        Loop While Not .atEndOfStream                       ' Continue until reach
                                                            ' the end of the CSV file.
    End With
    Set CSVint = Nothing
    Set StreamReader = Nothing
End Sub
```

So far, it has been outlined the way in which you can import the records from a CSV file sequentially, the following example shows how to filter the records, in a like SQL way, according to whether they meet a criterion set by the user:

```
Sub QueryCSV(path As String, ByVal keyIndex As Long, queryFilters As Variant)
    Dim CSVint As CSVinterface
    Dim CSVrecords As ECPArrayList
    
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"
        .dialect.fieldsDelimiter = ","                      ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf                  ' Rows delimiter
    End With
    If path <> vbNullString Then
        '@-----------------------------------------------
        ' The following instruction will filter the data
        ' on the keyIndex(th) field.
        Set CSVrecords = CSVint.GetCSVsubset(path, _
                                            queryFilters, _
                                            keyIndex)
        CSVint.DumpToSheet DataSource:=CSVrecords           ' Dump result to new WorkSheet
        Set CSVint = Nothing
        Set CSVrecords = Nothing
    End If
End Sub
```

In some situations, we may encounter a CSV file with a combination of `vbCrLf`, `vbCr` and `vbLf` as record delimiters. This can happen for many reasons, but the most common is by adding data to an existing CSV file without checking the configuration of the previously stored information. These cases will break the logic of many robust CSV parsers, including the demo of the 737K weekly downloaded [Papa Parse](https://www.papaparse.com/demo). The next example shows how an user can import CSV files with mixed line break as record delimiter, an option that uses the `multiEndOfLineCSV` property of the [`parseConfig`](https://ws-garcia.github.io/VBA-CSV-interface/api/properties/parseconf.html) object to work with these special CSV files.

```
Sub ImportMixedLineEndCSV()
    Dim CSVint As CSVinterface
            
    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Mixed Line Breaks.csv"
        .dialect.fieldsDelimiter = ","        ' Columns delimiter
        .dialect.recordsDelimiter = vbCrLf    ' Rows delimiter
        .multiEndOfLineCSV = True             ' All delimiters will be turned into vbLf
    End With
    With CSVint
        .ImportFromCSV .parseConfig
    End With
    Set CSVint = Nothing
End Sub
```

In all the above examples, an implicit assumption has been made, and that is that the user knows the configuration of the CSV file to be imported, so the question arises: can it be possible that the user does not know the configuration of the file to be imported? It is certainly possible, so how can the CSV interface help in these cases?

The tool includes a utility to sniff/guess field delimiters, record delimiters and quote character. This can be done with code like the following:

```
Sub DelimitersGuessing()
    Dim CSVint As CSVinterface

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = "C:\Sample.csv"           ' Full path to the file, including its extension.
    End With
    With CSVint
        Set .parseConfig.dialect = .SniffDelimiters(.parseConfig)    ' Try to guess delimiters and save to internal
                                                                     ' parser configuration object.
        '@--------------------------------------------------------------
        ' *NOTE: the user can also create a custom configuration object
        '        and try to guess the delimiter with it.
    End With
End Sub
```

## Contributing

In order to contribute within this project, please see the [guidance for contributing](https://ws-garcia.github.io/VBA-CSV-interface/contributing.html).

## Benchmark

The benchmark results for VBA-CSV Interface are available at [this site](https://ws-garcia.github.io/VBA-CSV-interface/home/getting_started.html#benchmark).

## Dependencies

The library depends on the [ECPTextStream class]( https://github.com/ws-garcia/ECPTextStream) in order to work with text files. In the same way, the class uses two external class modules: one for configuration sharing and the other for data storage. All dependencies are written in pure VBA.

## Limitations

Visit [this site](https://ws-garcia.github.io/VBA-CSV-interface/limitations/csv_file_size.html) in order to known around CSV file size considerations.

## Licence

Copyright (C) 2022  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

