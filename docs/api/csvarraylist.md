---
title: CSVArrayList
parent: API
nav_order: 5
---

# CSVArrayList
{: .d-inline-block }

New
{: .label .label-purple }

Class module developed to emulate some functionalities from the `ArrayList` present in some most modern languages. The `CSVArrayList` serve as a container for all the data read from CSV files and can be used to manipulate the stored items, or to store data that does not come from a CSV file, according to the user's request.
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
<td style="text-align: left; color:blue;"><em>Insert</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Inserts an Item, at the given Index, in the current instance of the class.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>IsJaggedArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Returns <code>True</code> if the paseed argument is a jagged array.</td>
</tr>
<tr>
<td style="text-align: left; color:blue;"><em>item</em></td>
<td style="text-align: left;">Property</td>
<td style="text-align: left;">Gets or sets an Item, by its index, from the current instance. This is the default property, so the user can use abbreviated expressions such as <code>expression(i)</code> to access the Item <code>i</code>, where <code>expression</code> represents a <code>CSVArrayList</code> object.</td>
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
<td style="text-align: left; color:blue;"><em>MultiDimensional</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Checks if an array has more than one dimension and returns <code>True</code> or <code>False</code>.</td>
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
<td style="text-align: left; color:blue;"><em>TwoDimToJaggedArray</em></td>
<td style="text-align: left;">Method</td>
<td style="text-align: left;">Turns a two-dimensional array into a jagged array. The method will successively deconstruct and delete the two-dimensional, passing its contents to the specified jagged array array.</td>
</tr>
</tbody>
</table>

[Back to API overview](https://ws-garcia.github.io/VBA-CSV-interface/api/)