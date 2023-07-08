
# Instruction for table_functions.xlsx

_Compatible with: Excel 365_

_Applies to file: https://github.com/ammosov/finmodel/blob/main/table_functions.xlsx_

If we store information in Excel tables, we need a method to get information out of tables and use it elsewhere. Excel has a built in structured reference language that can do a few things - 
for example, reference a table column as an array. Structured references have some limitations - for example, they behave like relative references when copied and cannot directly convert a text string from cell into a table or a column label.  

Before you begin, install add-on "Excel Labs" from Microsoft Store (former Advanced Formula Environment, see https://github.com/microsoft/advanced-formula-environment ). 
This feature was first published in 2023, and is a mini-IDE within Excel. Must have. We will use it to create and edit named functions. 

Named function encapsulate logic (i.e. operatons and calculations) that are repeated many times, have a lot of steps and (most important!) easy to break by a random typo. We will hide this complex logic in a container that will require only references to a few cells as user input. 

## Input

Open file `table_functions.xlsx`. It contains a sample table `Holidays` of 15 worldwide holidays that are huge retail events too. Each record has a unique number `[Id]`, holiday name as text string `[Holiday]`, holiday date as value `[Date]`, and holiday duration in days as integer `[Days]`. You do not need to make any changes to see how it works.

Click on Excel Labs icon to display IDE side panel and open Modules tab. You will see there the code for three custom functions that are used in this tutorial. 

![image](https://github.com/ammosov/finmodel/assets/4894284/b12ba0c9-4fae-4bf8-9646-a3b93d9fec4b)

## Output

We will retrieve all and any table data using a single interface. Use dropdown list in cell F2 to select a column label. 

![image](https://github.com/ammosov/finmodel/assets/4894284/955db5b9-a8c7-46bb-b3cd-385f832984b6)

The spreadsheet will display in cells F3:F17 an array of values from the selected column. 

![image](https://github.com/ammosov/finmodel/assets/4894284/58e3ccb3-183c-41dc-8c22-8723a5e3431c)

Use dropdown list in cell G2 to select a field label. The list in G2 will change every time when you make a change in F2 (but the previous value will remain even if invalid under a new list validation, and may display an error).   

![image](https://github.com/ammosov/finmodel/assets/4894284/04531c5e-0a65-4f08-8f3e-6d1ee390a598)

The spreadsheet will display in cells G3# an array of one or more rows. 

![image](https://github.com/ammosov/finmodel/assets/4894284/be0a6bab-d83d-4281-b08c-693d99c907b2)

Use parameters in cells G14, G15 and G16 to retrieve individual cells and blocks of cells. G14 and G16 contain dropdown list, G15 is a single string cell. Try typing different strings in G15 and see what happens. 

![image](https://github.com/ammosov/finmodel/assets/4894284/9dcb7164-dd5c-4b32-b2fd-8d5152b136f7)

All these techniques may be useful for creating Excel dashboards, both financial and not. What is more important, though, is that these functions return arrays that can be used inside calculations, and this is how we will reuse them further on. 


## Under the hood

### Excel notation for "range" and "array"
Range: a rectangular block of cells. 

Example: `A1:B2` <=> 2x2 block; row 1: ` | A1 | B1 |`, row 2: `| A2 | B2 |`

	A1 | B1
	A2 | B2 


Array: a matrix of several values used in calculation as one group. 

Example: `{1,2;10,20}` <=> 2x2 matrix; row 1: `| 1 | 2 |` ; row 2: `| 10 | 20 |`

	 1 |  2
	10 | 20 


### Table manipulation functions

	getTableCol = LAMBDA(
	    table,
	    column,
	    INDIRECT(
	        TEXT(table,"")
	        & "["
	        & TEXT(column,"")
	        & "]"
	        )
	);
 
	getTableRow = LAMBDA(
	    table,
	    id,
	    id_value,
	    FILTER(
	        INDIRECT(TEXT(table,"")),
	        getTableCol(table,id)=id_value
	    )
	);

	getTableCell = LAMBDA(
	    table,
	    field,
	    id,
	    id_value,
	    FILTER(
	        getTableCol(table,field),
	        getTableCol(table,id)=id_value
	        )
	);


- Table manipulation is done by filtering. 
- All functions return either a single value or an array of values.
- `getTableCol(table,column)` returns a full "table" column as array. It is no different from `Table[Column]` reference except it is a dynamic construction and can be changed through arguments in reference cell.  
- `getTableCell(table,field,id,id_value)` finds one or several cells in one "table" column ("field") marked by "id_value" in another column ("id"). 
- `getTableRow(table,id,id_value)` returns the whole "table" row as array if the "id" column has the "id_value". 
- Both `getTableCell()` and `getTableRow()` functions can either call a unique table record id ("UID") or be used to search the table by a key string.
- This set of functions uses one criteria only. Multiple criteria can be added, but the functions will be more convoluted.
- All arguments to all functions are either text string or references to cells with text strings. `getTableCol("Table1","Column1")` - OK. `getTableCol(A1,A2)` - OK IF `A1` contains "Table1" and `A2` contains "Column1". `getTableCol(Table1,Table1[Column1])` - not OK and will return an error!  
- `FILTER(range,criteria)` takes in two or more ranges of equal dimensions. Range 2 must be `TRUE-FALSE` or `0-1`, and the function selects from Range 1 the same lines than have `TRUE` or `1` values in Range 2.
- `INDIRECT()` takes a text string that is identical to a valid Excel range address and converts it to a valid Excel address. I am not sure why Excel developers did it this way (and suspect some need for compatibility with some obscure elements of Excel core as written in late 1980s).  
- `TEXT()` is a function intended specifically for converting its argument from number to text; here it provides an extra control for cases when the referring cell contains a number instead of a text string. Named structured references to tables are not converted by `TEXT()`.

### Data Validation Dynamic List 

Data Validation is an Excel feature that allows to create a drop down list in cell. As of 2024, Excel 365 can use only ranges in DataValidation, not formulas or names that return an array. Until a new version makes it possible, we need to use a workaround for this. One way is to use `INDIRECT()` to convert a dynamic reference to a finite range that data validation can use. 

`=INDIRECT("Table[#Headers]")` will create a dropdown list of "column" labels. 

`=INDIRECT("Table[Column]")` will create a list of values in a "column".

`=INDIRECT("Table["&A1&"]")` can read a column label from cell `A1`.

![image](https://github.com/ammosov/finmodel/assets/4894284/4d0e37c6-f125-4740-aedb-830a0914461c)

Note that the lists generated this way will contain unique values (as if `UNIQUE()` function was used on them) but cannot be sorted yet, they follow the sort order of the source. 

### Conditional Format for different data types

Excel stores dates as numbers, and it is not possible to carry the specific date format along with this number across calculation. Date as number, while perfectly correct, is hard to read and interpret. A crude but working workaround used here is conditional format that checks the range of dates acrually used in table, and then, if the number falls within this range, formats it as date. Of course, a potential problem is a different data type number that will be also within this range. However, as long as date as numbers are over 40 000 and ids and days are double digits at best, no conflict of data types is expected. 

Keep in mind: this format must by applied to all cells where the array can "spill", not just the upper left corner cell where the array generator formula is located. 

`=AND($F$3>=MIN(INDIRECT("Holidays[Date]")),$F$3<=MAX(INDIRECT("Holidays[Date]")))`

![image](https://github.com/ammosov/finmodel/assets/4894284/4934e6e8-ca0b-40c1-9dde-e463d8997d1c)

Note: array for getTableCell() in `I16#` is a different rule that checks value of `I14`, and uses a different date format (verbose month names). 

![image](https://github.com/ammosov/finmodel/assets/4894284/42f5f52c-067e-4c8a-8e66-2a20f03e58a9)

### Excel legacy bugs and errors

For a very detailed discussion of how and why Excel for many years used two different methods (with different results) 
for calculating linear regreression for its own `LINEST()` function and for making trendlines in charts, see: 

https://www.informit.com/articles/article.aspx?p=2019170
