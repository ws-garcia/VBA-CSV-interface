---
title: CSVArrayList
parent: API
nav_order: 5
---

# CSVArrayList
{: .fs-6 }

Class module developed to emulate some functionalities from the `ArrayList` present in some most modern languages. The `CSVArrayList` serves as a container for all the data read from CSV files and can be used to manipulate the stored items, or to store data that does not come from a CSV file, according to the user's request.
{: .fs-4 .fw-300 }

---

## Members

<table>
<thead>
<tr>
<th style="text-align: left;">Item</th>
<th style="text-align: left;">Type</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left; color:blue;"><em>Add</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Appends a copy of the specified value to the current instance. Given the nature adopted by the CSV interface to store data, if the value to be appended to the current instance is not a one-dimensional array, where each element represents a field, the user will not be able to use data sorting methods properly. User must use the <code>Add2</code> method instead if the goal is to sort stored items.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Add2</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Appends a copy of the specified values to the current instance. In contrast to the <code>Add</code> method, the data is operated on before being stored, so if the values to be appended to the current instance are not one-dimensional arrays, they will be properly stored as one-dimensional array. In this way, the user will be able to use the data sorting methods provided by the class as long as no multi-dimensional arrays are stored in the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>AddIndexedItem</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Appends a copy of the specified values to the current instance using a string-type key. This allows access to the elements by providing an index or a key. If the key exist, the item will be modified only if the <code>UpdateExistingItems</code> parameter is ser to <code>True</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Clear</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reinitializes the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Clone</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns a <code>CSVArraylist</code> as a exact copy of the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Concat</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Concatenates the values from the current instance with the specified values and returns a <code>CSVArraylist</code> object as result. The <code>AValues</code> parameter is a <code>Variant</code> data type containing the array, <code>CSVArraylist</code> or value to concatenate.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Concat2</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Concatenates the values from the current instance with the specified values and returns a <code>CSVArraylist</code> object as result. The <code>AValues</code> parameter is a <code>CSVArraylist</code> with the values to concatenate.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Copy</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns a <code>CSVArraylist</code> object with a copy of the current instance from and to a given index. The <code>StartIndex</code> parameter indicates where the copy will start and the <code>EndIndex</code> determines where the operation will end. If the <code>EndIndex</code> parameter is set to <code>-1</code>, the operation will end at the maximum index available for the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>CopyToArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns an array with a copy of the current instance from and to a given index. The <code>StartIndex</code> parameter indicates where the copy will start and the <code>EndIndex</code> determines where the operation will end. If the <code>EndIndex</code> parameter is set to <code>-1</code>, the operation will end at the maximum index available for the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>count</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Returns the amount of items stored in the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>CreateJagged</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Creates an empty jagged array. The operation will turns the array <code>ArrVar</code> into an jagged array with <code>ArraySize + 1</code> rows and each row with <code>VectorSize</code> columns. To access to an individual element user must use something like <code>expression(i)(j)</code>, where <code>i</code> denotes an index in the main array and <code>j</code> denotes an index in the child array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Dedupe</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Remove duplicates from records. Requires rectangular table input (all records with same fields count).The <code>keys</code> parameter will indicate which fields/columns will be used in the deduplication. A string like "0,5" used as parameter will deduplicate the records over columns 0 and 5. If a string like "1-6" is used as argument, the deduplication will use the 2nd through 7th fields.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Filter</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns a filtered array list using the <code>CSVexpressions</code> class module.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>FromString</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Main constructor method. Populates the current instance using values passed as a Java array string (<code>\{\{*\};\{*\}\}</code>).</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>GetIndexedItem</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Gets an indexed Item, by its key, from the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Group</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Groups rows having the same values into a summary.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>IndexedItems</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets all indexed Items from the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>indexing</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Indicates whether the current instance is used to store indexed elements.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Inner, Left and Right Join</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Run a like SQL join on the provided data tables.<br>1) Use a string such as <code>{1-2,5,ID};{1-6}</code> as a predicate of the columns to indicate the join of columns 1 to 2, 5 and ID of leftTable with the columns 1 to 6 of rightTable.<br>2) Use a string such as <code>{*};{1-3}</code> to indicate the union of ALL columns of leftTable with columns 1 to 3 of rightTable.<br>3) The predicate must use the dot syntax <code>[t1.#][t1.fieldName]</code> to indicate the fields of the table, where t1 refers to the leftTable.<br>4) The matchKeys predicate must be given as <code>#/$;#/$</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Insert</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Inserts an Item, at the given Index, in the current instance of the class.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>InsertField</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Inserts a new field named <code>FieldName</code> into the records of the current instance at the given index. If a formula is provided, the field is populated in each record (row) with the result of evaluating the formula using the fields specified in the formula.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>IsJaggedArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns <code>True</code> if the paseed argument is a jagged array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>isSorted</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Returns <code>True</code> if the current instance data is sorted.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>item</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets an Item, by its index, from the current instance. This is the default property, so the user can use abbreviated expressions such as <code>expression(i)</code> to access the Item <code>i</code>, where <code>expression</code> represents a <code>CSVArrayList</code> object.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ItemExist</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Checks if a given field exists in a record of the current instance. Returns <code>False</code> when the key can not be found.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ItemIndex</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Performs a search, on a given field, and retrieves the index of the target record. USE ONLY WITH SORTED DATA.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ItemKey</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Gets the key at given position.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>items</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets the collection of elements from or to the current instance. To set the elements, the <code>AValue</code> parameter must be an array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>JaggedToTwoDimArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Turns a jagged array into a two dim array. The method will successively deconstruct and delete the jagged array, passing its contents to the specified two-dimensional array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>KeyExist</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Searches for a key in the internal indexed records and returns <code>True</code> when found.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>KeyIndex</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Searches for an element in the internal indexed records, using a key, in the current instance (ONLY when the data is already sorted in ascending order). Returns the index of the element when found and -1 when the key is not found.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Keys</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets all keys for all indexed Items from the current instance.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>keyTree</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Allows to store and group elements with the same key.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>lastSortedIndex</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Retrieves the index of the field used in the last sort operation.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>MergeFields</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Merges the specified fields in the current instance data table. The <code>indexes</code> parameter will indicate which fields/columns will be merged. A string like "2,7" used as parameter will merge the records over the columns with indexes 2 and 7. If a string like "3-8,10" is used as argument, the merge operation will use the 4th to 9th fields and the 11th field.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>MultiDimensional</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Checks if an array has more than one dimension and returns <code>True</code> or <code>False</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RearrangeFields</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Rearranges the fields of the stored data. A string such as "0-3,5-4,6-11" used as a parameter will leave the position of fields with indexes 0 to 3 unchanged, swap the fields at indexes 5 and 4, and leave all remaining fields in position.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Reduce</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reduces the internal array list to the result by evaluate the <code>ReductionExpression</code> parameter over all items.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Reinitialize</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reinitializes the current instance of the class and reserves the storage space desired by the user through the <code>bufferSize</code> parameter.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RemoveAt</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Removes the Item at specified Index.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RemoveField</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Removes the field at specified <code>aIndex</code> in all records.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RemoveIndexedItem</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Removes an indexed Item using the specified key.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>RemoveRange</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Removes a range of Items starting at the specified Index.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Reverse</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reverse the order of the internal items, from a given <code>StartIndex</code> to a <code>EndIndex</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Reverse2</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Reverse the order in the target jagged array, from a given <code>StartIndex</code> to a <code>EndIndex</code>.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ShiftField</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Moves a field leftward or rightward. Negative values for the <code>Shift</code> argument will produce leftward shifts.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ShiftRecord</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Moves a record upward or downward. Negative values for the <code>Shift</code> argument will produce upward shifts.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ShrinkBuffer</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Shrinks the buffer size to avoid extra space reservation.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Sort</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Sorts the internal Items. Requires all Items to be one-dimensional arrays. If the <code>FromIndex</code> is set to <code>-1</code>, the sorting will start at the Items lower bound; when the <code>ToIndex</code> is set to <code>-1</code>, the operation will end at the Items upper bound. The <code>SortingKeys</code> parameter is used to define the index of the columns on which the sorting operation will be performed, negative values indicate sorting in descending order; the user can pass an array of sorting keys as a parameter. The <code>SortAlgorithm</code> parameter indicates which sort algorithm will be used to perform the sort.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>SortByField</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Sorts the internal items by an specified field.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>SortKeys</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Sorts the internal Items by its keys using QuickSort. Requires all Items to be one-dimensional arrays. The indexes are base 0.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>SplitField</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Splits the specified field in the current instance data table.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Swap</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Swap Items in buffer.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>Swap2</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Swap Items in target jagged array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>ToString</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Serializes the buffer contents to a common string representation. Only one-dimensional arrays and jagged arrays populated with one-dimensional arrays are supported.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>TwoDimToJaggedArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Turns a two-dimensional array into a jagged array. The method will successively deconstruct and delete the two-dimensional, passing its contents to the specified jagged array array.</td>
</tr>
</tbody>
</table>

[Back to API overview](https://ws-garcia.github.io/VBA-CSV-interface/api/)
