'The following Script will iterate through an excel spreadsheet containing data for stocks
'for each trading day within a year, for numerous stocks.'

'The data includes daily price datapoints for;
'opening, closing, highest and lowest.

'There will be a columnn describing each datapoint type'

    'The script will extract data for each stock and place them in a summary table;


'This data will include:Yearly Change, Percent Change and Total Stock Volume.
            

	'Yearly change from the opening price at the beginning of a given year to the 	closing price at the end of that year.

            
	'The percentage change from the opening price at the beginning of a given year 	to the closing price at the end of that year.


	'The total stock volume of the stock.

The script will then extract the following values from the summary table:
            'greatest percentage increase & Corresponding Ticker Symbol
            'greatest percentage decrease & Corresponding TIcker Symbol
            'greatest volume & corresponding Ticker Symbol'


LOGIC:

Lucky for us, The data is arranged in chronological order and by stock symbol:

Therefore we know:
	The first rows correspond to an entire dataset of stocks before changing to 	the next stock.
	Any givens stocks first and last row of datapoints will correspond to 	the 	first and last day of opening/closing price of that year.

Thus:
	'we need the script to identify, for any given stock, when first row of data 	 and when the last row of data is being analysed.'	
	
	'We will use a boolean to tell us that the row is first in each stocks 	dataset'
        
Using If Statements:
	'To identify the last row of a stock within a block, the next row will not 	equal the current row'
		Resets variables being used to store information, calculates and 		places them within the summary table.
		we know the next iteration will be the first row for the next stock, 		so reset boolean first row to equal True.
	
	' Determine if the current row i is equal to the next row, i.e. not the last 	row within a block of a stock, checks if the Boolean identifying the first 	row is True. If so, takes the opening value of that row.

The script will then extract the following values from the summary table:
            'greatest percentage increase & Corresponding Ticker Symbol'
            'greatest percentage decrease & Corresponding TIcker Symbol
            'greatest volume & corresponding Ticker Symbol'

To do this we iterate through the summary table, storing the value required then extracting them at the end.

IF Statements:
	'As loop iterates through find the greatest value,
	IF value found that is greater than the value currently stored, it must then 	be stored instead.

	'As loop iterates through find the lowest value,
	'IF the value found that is less than the Value currently stored, it must 	then be stored instead'



	