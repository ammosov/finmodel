# Instruction for monthly_salaries.xls

_Compatible with: Excel 365_

_Applies to file: https://github.com/ammosov/finmodel/blob/main/monthly_salaries.xlsx_

Let us begin with some of the simplest tasks, like making a monthly salary payment schedule.  

## Input
Go to `data` tab. I created there a sample Excel table of eight employees. Excel table is a smart data structure that extends itself when data is added to it (but does not contract when data is deleted). 

Type the data for your first 8 employees over the 8 sample employees. You will need a position name (Founder, CTO, Developer etc), salary sum, start and end dates. Once you go to Employee 9, just type `9` in `id` column, and the new row will be added automatically. 

End date is needed to denote temporary employees or contractors, such as a Lawyer in the example who works for 3 initial months. For permanent positions, set a larger than life date, like in year 2099. If you have less than 8 employees planned, do not forget to delete the rows manually - it will be the only change you will do manually. 

![image](https://github.com/ammosov/finmodel/assets/4894284/f8c114cd-a0d9-48a7-8af4-30d3d6ef512a)

## Output
Go to `Monthly` tab. Type in the model parameters: 
- In cell B1 the year in which your financial model starts;
- In cell B2 the number of the starting month (January=1, February=2...); 
- In cell B3 the number of months for which the financial model should last.

The spreadsheet will auto generate a schedule of salary payments for each employees by month, together with a sum of total salary payments.   

![image](https://github.com/ammosov/finmodel/assets/4894284/3b6e3cfd-1b22-49e8-8998-952393e024f7)

For those of you who care what happens under the hood:

### Create a flexible date sequence in a row

`=DATE(year_start,SEQUENCE(1,month_start,how_many_months),1)`

First we take a DATE() function and feed to it an array of month numbers to force it to convert it to an array of dates
- `DATE(year,month,day)` creates a date value for a given year, day and month
- `year`=year_start is the year when the sequence starts, e.g. 2023;
- `month`=SEQUENCE() is an Excel function that generates a single row of month numbers to be converted into dates;
- `day`=1 tells DATE() that each date will be the 1st day of the month.

`SEQUENCE(rows,columns,start,step)` returns an array of x `rows` by y `columns`, beginning from `start` value and incrementing it by `step` value; if `step` is omitted, it defaults to 1, as in our case. 

Example used here at C5: `=DATE($B$1,SEQUENCE(1,$B$3,$B$2),1)`

### Return a 1-dimension array of sums of columns of another array

`=BYCOL(start_address#,LAMBDA(arg*,FUNCTION(arg*)))`
- `BYCOL (array, function)` is a function that tells Excel to take an array and apply the function to it column by column;
- `array`=start_address# is the `address of start (upper left) cell` of the array to be summed and followed by `#`.
- `#` (hash) is a "spill operator" that tells the function: "this `array` is "dynamic" and has a variable length";
- `LAMBDA()` is a wrapper that passes `arg*` (i.e. one of more arguments) to the `FUNCTION` and returns result; arguments always come first, function last.

Example used here at C6: `=BYCOL(C8#,LAMBDA(a,SUM(a)))`  

### Return a table column as single vertical array

`=Table_name[column_name]`

Example used here at B8: `=Employees[position]`

### Get an array of sums to be paid in between these days

This is where things get tricky! We begin with creating an invisible boolean 0-1 array that answers the question "Is a given person employed after one date and before another", where 1=employed and 0=not employed. Since both conditions must be met simultaneously (AND() operator), we multiply two arrays that answer two questions: "Is this date on or after known date when employment starts" and "Is this date before known date when employment ends". Then we multiply the invisible array of 1s and 0s by this person's salary, which results in a final visible dynamic array of salary payments that matches the dynamic list of employees and the dynamic list of monthly periods.  

    `=(Employees[salary_sum])*
    
    --(DATE($B$1,SEQUENCE(1,$B$3,$B$2),1)>=Employees[start_date])*
    
    --(DATE($B$1,SEQUENCE(1,$B$3,$B$2),1)<Employees[end_date])`

- `--()` is a "double unary" operator that force converts `TRUE` and `FALSE` bools to `1` and `0` values (TRUE > -1 > 1 and FALSE > 0 > 0). In Excel, this type conversion needs to be done explicitly. Another way to do it is to use `N()` formula.   
