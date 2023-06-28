
# Instruction for monthly_journal.xlsx

_Compatible with: Excel 365_

_Applies to file: [https://github.com/ammosov/finmodel/blob/main/monthly_journal.xlsx](https://github.com/ammosov/finmodel/blob/main/monthly_journal.xlsx)_

An important part of financial modeling is mapping sums of money to periods and categories. The salary forecast exercise did it in a most simple form, because salary payments are the same every time and happen once every period. Let us deal with more complicated data that consists of individual payments on separate dates. Payment can be incoming, such as invoices paid by startup customers, and outgoing, such as cloud service bills, conference fees, legal expenses, and nearly everything else.  

We will summarize a "journal of payments" by periods and categories.    


## Input

Go to `data` tab. I created there a sample Excel table of sixteen payments for typical startup services. One is rent, another is cloud service, and two are subscription services. All of them are paid at more or less the same time of the month, but not on the same day. Sums can be different, too. All of the sums in out example are negative which means they are costs, that is, paid by a startup to someone else. Such a list of payments is called "journal". 

![image](https://github.com/ammosov/finmodel/assets/4894284/67712314-0d39-4eb6-bec2-1dd6d04798e9)

You can type your own data in this table. Note that there are only three categories that you can use: _"cloud"_, _"subscription"_ and _"rent"_. If you need some other category, go to `lists` tab and add a category in Column A below the existing list. New items will be automatically added to the drop down list in table and sorted alphabetically. 

![image](https://github.com/ammosov/finmodel/assets/4894284/3f68995f-a87b-433d-ad79-51b13736cf24)

Also on `lists` tab, type in the model parameters: 
- `date_start` is a date on which your model starts; it can be any date
- `date_period` is a "time step" of your model in months; every period will start on the same day of month as your model start date
- `date_end` is the number of periods in the model.

![image](https://github.com/ammosov/finmodel/assets/4894284/11ade8be-6db9-4d6a-8856-b940269e94a4)

## Output

Go to `Monthly` tab. The spreadsheet will auto generate a summary table. The table will have sums of all payments by each category (row) and for each month (column).

![image](https://github.com/ammosov/finmodel/assets/4894284/5cd5a248-efb0-4dcc-ad0a-98b4358cc944)	

## Under the hood

### Named cells and cell ranges

Names are stored in Name Manager and are labels that refer to cells, ranges, expressions and formulas. Names are a convenient way to reuse the code, that reduces mistakes and make the code much easier to read. 

In this example, we defined three variables and two expressions `categories` and `date_sequence`. 

`date_start` = lists!$D$3

`date_period` = lists!$D$4

`date_end` = lists!$D$5

### Dynamic list of categories

	categories = 
		SORT(
			OFFSET(
				lists!$A$2,
				0,
				0,
				COUNTA(lists!$A$2:$A$999)
				,1
			)
		)

- `OFFSET()` function takes an address of a start cell, rules for getting to end cell and return an address of a rectangular range between start and end cells
- `COUNTA()` function counts all nonblank cells in a very large range (up to 999 cells in this example), asserting they are all adjacent, and returns the count as length of non-empty range to `OFFSET()`
- `SORT()` function sorts a range, alphabetically by default 

Result is a dynamic list of categories that can be called as `=categories` or used in formulas. 

### Named sequence of dates

	date_sequence = 
		DATE(
			YEAR(lists!$D$3),
			SEQUENCE(
				1,
				date_end,
				date_period,
				1
			),
			DAY(date_start)
		)

- `DATE(year,month,day)` is a function that creates a date value from number for a year, month and day
- `YEAR(datevalue)`, `MONTH(datevalue)` and `DAY(datevalue)` extract year, month and day from date as value
- Instead of giving a `DATE()` function a single number for month, we give it a range of numbers `(SEQUENCE(date_end, date_period)` and force it to return a range of dates one date_period apart for date_end number of periods

### Dynamic summary table

This expression is quite complicated and is essentially similar to list comprehension and for loops in Python. It is used only to create a matrix that is flexible in two dimensions and is fully automatic. If such a matrix is to be created manually, a number of much simpler methods are available. In particular, `SUMPRODUCT()` can be easily used for matrix manipulations. 

The summary table expression here has one advantage: it can take any number of periods and categories and return a correct range of sums. 

	=MAKEARRAY(
		COUNTA(categories),
		COUNT(date_sequence),
		LAMBDA(
			a,
			b,
			SUM(
				FILTER(
					Journal[Sum],
					--(Journal[Category]=INDEX(categories,a)
					)*
					--(Journal[Date]>=INDEX(date_sequence,b)
					)*
					--(Journal[Date]<DATE
						(
							YEAR(
								INDEX(date_sequence,b)
							),
							MONTH(
								INDEX(date_sequence,b)
							)+date_period,
							DAY(
								INDEX(date_sequence,b)
							)
						)
					)
				,0)
			)
		)
	)

- `MAKEARRAY(x,y,expression)` creates an `x` by `y` dynamic array and fills it with results of expression
- `COUNTA()` counts the number of nonblank cells in a range - which is, in our case, the same as length of `categories` and the number of rows in our summary matrix; we use this function because text strings cannot be counted directly
- `COUNT()` counts the number of values in a range - which is, in our case, the same as length of `date_sequence` and the number of columns in our summary matrix
- `LAMBDA(a,b,expression(a,b))` inside MAKEARRAY() takes the matrix height `a` and width `b` we just calculated, iterates over them and sends the value pairs to the `expression(a,b)`
- `FILTER(range,criteria)` returns a subrange of `range` that meets `criteria`; SUM(FILTER()) sums this subrange and returns a single summary value
- Criteria inside `FILTER()` are 0-1 boolean ranges of the same length; multiplication serves as `AND()` operator
- `INDEX(range,x)` is a function that returns `xth` element from `range`
- `0` in the last position of `FILTER()` is a substitute for "no journal record is found for this period or category" 

When this function runs, the following happens: we take a period-category pair, check which sums in the journal have the date in between the period's start and end dates and are labeled with the category, sum all the sums we found, place it in the correct place of the matrix - and repeat with next paid, until all possible pairs are found and checked.    
