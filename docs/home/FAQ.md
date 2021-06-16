---
layout: default
title: FAQ
parent: Home
nav_order: 5
---

<center> <h1>FAQ</h1> </center>

<center> <h2>General</h2> </center>

### Why use VBA CSV interface?

Many solutions have been developed to solve the problem of working with CSV files, some enthusiasts have worked the problem in VBA, an action that is considered by many as unsuccessful and unfortunate since Excel has tools such as Power Query that allow importing data into worksheets in a simple way and with great processing power.

However, using Power Query implies establishing a physical connection to each data source (CSV file). If you want to process the data from VBA, using data arrays to store the imported information (without create a connection, write data to a worksheet, copy the information to an array, delete the connection and delete the data from the worksheet) you must use VBA CSV interface to achieve your goals.

### Does VBA CSV interface have any dependencies?

Yes. The VBA CSV interface dependencies are:  

* parserConfig.cls
* ECPTextStream.cls
* ECPArrayList.cls

Please review the [instalation page](https://ws-garcia.github.io/VBA-CSV-interface/home/installation.html). 

### Which Microsoft applications is it compatible with?

For the moment, VBA CSV interface can run over Microsoft Excel and Microsoft Access. 

### Can I contribute something?

Yes, Please! VBA CSV interface is an open source project. I don't want to do this all by myself. Take a look to the [GitHub project page](https://github.com/ws-garcia/VBA-CSV-interface) and hack the code. If you're making a significant change, open an issue first so we can talk about it. 

<center> <h2>Performance</h2> </center>

### Why do recent versions of VBA CSV interface have lower performance?

Some users of the VBA CSV interface have complained that the library has inferior performance to the prior versions. It should be noted that, starting with version 3 of this library, the choice has been made to offer [greater usability at the cost of performance loss](https://ws-garcia.github.io/VBA-CSV-interface/home/getting_started.html#philosophy).

For example, it has been decided to use jagged arrays and a specialized class module to store the information imported from CSV files. How does this decision impact performance? In the early versions of VBA CSV interface the data was stored in two-dimensional arrays of type `Strings`, since version 3 the `Variant` data type is used which [can be translated into performance loss](https://www.aivosto.com/articles/stringopt.html). In the same vein, the implementation of [jagged arrays results in a loss of performance in VBA](https://excelvirtuoso.wordpress.com/2018/08/13/jagged-arrays-vba/).

### Why the change from String data types to Variant type? 

Jagged arrays require the parent array to be of type `Variant` to contain another array inside, the inner arrays need to be declared as of type `Variant` to provide the user with dynamic typing capability, a feature not present in versions prior to VBA CSV interface v3.

### Why implement jagged arrays? 

Previous versions of VBA CSV interface appended additional empty fields to ALL records, this happened when importing CSV files in which the records did not have the same amount of fields in the whole file. As a result, for example, if the user decided to export the imported information from a file in which the first line (header) had 12 fields, with most of the records in this CSV only containing the first 5 or 7 fields, the file that was written (after exportation) contained several consecutive commas until the 12 fields per line were completed. This problem is solved with the implementation of jagged arrays.

In addition, the jagged arrays make it easier to implement the [Yaroslavskiy Dual-Pivot Quicksort algorithm](https://web.archive.org/web/20151002230717/http://iaroslavski.narod.ru/quicksort/DualPivotQuicksort.pdf), a feature that can be very useful in certain circumstances.

### Why sacrifice performance?

A good piece of code is not only the one that runs in the shortest time, but also factors such as usability and the flexibility with which it can be used. Previous versions of the VBA CSV interface ran at surprising speed, but had certain limitations that made them look rigid or inflexible. Whereas recent versions seek that balance between power, performance, usability and flexibility. 

For example, instead of offering a single way to get things done, the VBA CSV interface user can use the library's modules to achieve the goal in a variety of ways:

#### Import CSV files

1. Built-in methods: [ImportFromCSV](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsv.html) (import selected records or all records contained in the CSV file) and [GetRecord](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/getrecord.html) (read records sequentially, one after another, from the first to the last).
2. Alternative method: use the [ECPTextStream](https://ws-garcia.github.io/ECPTextStream/) module, with the [endStreamOnLineBreak](https://ws-garcia.github.io/ECPTextStream/api/properties/endstreamonlinebreak.html) property set to `True`, read text streams sequentially with the [ReadText](https://ws-garcia.github.io/ECPTextStream/api/methods/readtext.html) method and parse the text string stored in the [bufferString](https://ws-garcia.github.io/ECPTextStream/api/properties/bufferstring.html) using the [ImportFromCSVstring](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/importfromcsvstring.html) method.

Another limitation that has been broken in the most recent versions of the VBA CSV interface is linked to the size of the parsed file. In versions prior to v3, the performance was subject to the size of the file, due to memory usage, even if the user only needed to access the first few records of the CSV file. Thus, importing the first 100 records from a 200 MB file took much longer than importing the same number of records from a 10 MB file. In recent versions the performance is strictly linked to the number of fields and records being processed, since the entire contents of the CSV are NOT loaded into RAM thanks to the [ECPTextStream](https://ws-garcia.github.io/ECPTextStream/) module.

