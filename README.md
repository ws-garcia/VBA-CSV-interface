# VBA-CSV file Manager
## Table of contents
* [Intro](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#intro)
* [Advantages](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#advantages)
* [Philosophy](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#philosophy)
* [Rules](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#rules)
* [Usage](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#usage)
* [Benchmark](https://github.com/ws-garcia/VBA-CSV-fileManager/blob/master/README.md#benchmark)
## Intro
The CSV, stands from Comma Separated Values, files are special kind of tabulated plain text data widely used in data exchange. There is no globally accepted standard format for that kind of files, however, out there are well formed standards such as [RFC4180](https://www.ietf.org/rfc/rfc4180.txt) proposed by The Internet Society.
Although many solutions has been developed for work with CSV files into VBA, including projects from [@sdkn104](https://github.com/sdkn104/VBA-CSV) and [@Senipah](https://github.com/Senipah/VBA-Better-Array/tree/master/src) on Github, the vast majority of these have serious performance lacks. This argumentations conduce to the development of a VBA class module that allows users exchange data between VBA arrays and CSV files at relatively high speed.
### Advantages
* CSVfileManager can work with like RFC4180 CSV standard (there are few differences).
* User have the option to import only certain range of records from given CSV file.
* Writes and reads files at relatively high speed (tested against the available solutions).  
## Philosophy
The VBA CSVfileManager class module is designed for gain advantage from the well structured CSV files, this means, there isn't automatic syntax check, given the user decide how the class will works. This can be seen as a weakness, but the class get a speed-up on writing and reading procedures at time the user controls how the file is interpreted, keeping in mind that, in fact, VBA is a language with slow code execution speed. 
Under this idealization it's easy to develop a solution that implicity complies with the RFC4180 standart for user specified CSV document format. In order to achieve this, the user must to follow the rules specified below.
## Rules
1. Each record is located on a separate line, delimited by a line break (CRLF, CR, LF).
2. The last record in the file may or may not have an ending line break.
3. There maybe an optional header line appearing as the first line of the file with the same format as normal record lines.  This header will contain names corresponding to the fields in the file and should contain the same number of fields as the records in the rest of the file.
4. Within the header and each record, there may be one or more fields, separated by the fields separator (Comma, Semicolon, Space, Tab).  Each line should contain the same number of fields throughout the file.  **_Avoid let spaces betwen fields and records separators_**.  The last field in the record must not be followed by a fields separator.
5. Each field may or may not be enclosed in the selected Coerce char. **_The user can choose between enclose, coerce, every fields or neither one_**.
6. Fields containing special chars (line breaks, double quotes, apostrophe, and commas) should be enclosed with selected Coerce char.
## Usage
Import whole CSV file into an VBA array
```vbscript
Dim CSVmanager As CSVfileManager
Dim MyArray As variant
Set CSVmanager = New CSVfileManager
Call CSVmanager.OpenConnection(fileName)
Call CSVmanager.ImportFromCSV
MyArray  = CSVmanager .CSVdata
Set CSVmanager = Nothing
```
Import a range of records from CSV file into a VBA array
```vbscript
Dim CSVmanager As CSVfileManager
Dim MyArray As variant
Set CSVmanager = New CSVfileManager
CSVmanager.StartingRecord = 10
CSVmanager.EndingRecord = 20
Call CSVmanager.OpenConnection(fileName)
Call CSVmanager.ImportFromCSV
MyArray  = CSVmanager .CSVdata
Set CSVmanager = Nothing
```
Set the char to encapsulate, coerce, fields
```vbscript
CSVmanager.CoerceType = NotCoerce
CSVmanager.CoerceType = Apostrophe
CSVmanager.CoerceType = DoubleQuotes
```
Set fields and records delimiters
```vbscript
CSVmanager.FieldsDelimiter = ";"
CSVmanager.RecordsDelimiter = vbCrLf
```
## Benchmark
The class was tested against many solutions using the oldest, lowest-processing capacity laptop I could find: Win 7 Starter 32-bit, Intel® Atom™ CPU N2600 @1.60 GHz, 1 GB RAM. 
The times showed, seconds, in the bellow table are the average of ten (10) calls to the import procedure (supposed most costly to the CPU). The files used in the test haven twelve fields with variable number of records. 

|*Procedure (Author)*|*1K rec (102 KB)*|*5K rec (511 KB)*|*10K rec (0.99 MB)*|*100K rec (9.95 MB)*|
|:--------------------------|-----------------:|----------------:|----------------:|-----------------:|
|*ImportFromCSV (W. García)*|_0.0781_|_0.3021_|_0.4792_|_4.9141_|
|*ParseCSVToArray/ADO (@sdkn104)*|1.4349|47.3177|202.82|>1,000|
|*ImportCSVinArray (Wester)*|0.1042|0.6484|1.0182|10.250|
|*ArrayFromCSV (Heffernan)*|0.2396|1.7839|2.2057|22.385|
|*FromCSV(@Senipah)*|0.3594|3.8333|16.6172|>1,000|

Considering the system specification for the test machine (4 MB/sec. when it writes files to an USB), the above times was stunning!: up to 2.15 MB/sec. for reading operations.
