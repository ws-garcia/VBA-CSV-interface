---
title: tJoin
parent: Methods
grand_parent: API
nav_order: 31
---

# tJoin
{: .d-inline-block }

New
{: .label .label-purple }

Run a left, right outer or inner join on the provided data tables.
{: .fs-4 .fw-300 }

---

## Syntax

*expression*.`tJoin`*(nType, leftTable, rightTable, Columns, matchKeys, \[predicate= vbNullString\], \[headers=True\])*

### Parameters

<table>
<thead>
<tr>
<th style="text-align: left;">Part</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><em>nType</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>JoinType</code> enumeration variable representing the join nature.</td>
</tr>
<tr>
<td style="text-align: left;"><em>leftTable</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>CSVArrayList</code> Type object. Represents the first table in the join operation.</td>
</tr>
<tr>
<td style="text-align: left;"><em>rightTable</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>CSVArrayList</code> Type object. Represents the second table in the join operation.</td>
</tr>
<tr>
<td style="text-align: left;"><em>Columns</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Specifies the structure of the rows returned.</td>
</tr>
<tr>
<td style="text-align: left;"><em>matchKeys</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Represents the primary and preference keys to be matched.</td>
</tr>
<tr>
<td style="text-align: left;"><em>predicate</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>String</code> Type variable. Represents the condition that must be met when selecting rows.</td>
</tr>
<tr>
<td style="text-align: left;"><em>headers</em></td>
<td style="text-align: left;">Required. Identifier specifying a <code>Boolean</code> Type variable. Indicates if the tables have headers.</td>
</tr>
</tbody>
</table>

### Returns value

*Type*: `CSVArrayList`

---

See also
: [Filter method](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/filter.html).

## Behavior

Use a string such as "{1-2,5,ID};{1-6}" as a predicate of the columns to indicate the join of columns 1 to 2, 5 and 'ID' of leftTable with the columns 1 to 6 of rightTable. Use a string such as "{*};{1-3}" to indicate the union of ALL columns of leftTable with columns 1 to 3 of rightTable. The predicate must use the dot syntax [t1.#][t1.fieldName] to indicate the fields of the table, where t1 refers to the leftTable. The matchKeys predicate must be given as "#/$;#/$". 

### ☕Example

```vb
Sub Join(ByRef lTable As CSVArrayList, ByRef rTable As CSVArrayList)
    Dim CSVint As CSVinterface
    
    Set CSVint = New CSVinterface
    With CSVint
        ' Performs a Left join returning the "1st" and "Country" fields of the left table and the 
		  ' "Total_Revenue" field of the right table, joined in the "Order_ID" field of both tables, 
		  ' of those records that satisfy the given condition.
        .tJoin JoinType.JT_LeftJoin _
                lTable, rTable, _
                "{1,Country};{Total_Revenue}", _
                "Order_ID;Order_ID", _
                "t2.Total_Revenue>3000000 & t1.Region='Central America and the Caribbean'"
    End With
    Set CSVint = Nothing
End Sub
```

[Back to Methods overview](https://ws-garcia.github.io/VBA-CSV-interface/api/methods/)
