Attribute VB_Name = "Stock_SUMMARY_TABLE"
Option Explicit

Sub Stock_Summarizer()
    'Create a variable that allows us to iterate through the worksheets'
    'WS - Current Worksheet variable'
    Dim ws As Worksheet
    
    'Loop through the worksheets with WS - Representing the current worksheet'
    For Each ws In Worksheets
        

'The following Script will iterate through an excel spreadsheet containing data for stocks
'for each trading day within a year, for numerous stocks.'

'The data includes daily price datapoints for;
'opening, closing, highest and lowest.

'There will be a columnn describing each datapoint type'

    'The script will extract data for each stock and place them in a summary table,
    'This data will include:
            'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

            'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
                    
                    'For this we need, the opening price at the beggining of the year and the closing price at the end of the year,
                    'we then need to calculate the difference over the whole year, and the percentage of that change with respect to
                    'the opening price of that year'

            'The total stock volume of the stock.
                    ' for this we need to sum the trading volume of each day for each stock throughout the whole year'
                
            'A summary table for the extracted data to be stored within,
                    'as the summary table will not have as many rows as the whole dataset we need an index (counter),
                    'to put the summary of data into the table'
        
        
        
        'Define a variable to store:
        
        'The name of the stock Ticker           - The Ticker symbol of the stock being analysed'
        Dim Ticker As String
        
        'The TotalVolume of the stock           - The Volume sum of the stock through that stock's whole dataset'
        Dim TotalVolume As Double
        
        'The yearly open price of the stock     - the opening price on the first trading day of that year for each stock'
        Dim yearly_open As Double
        
        'The yearly closing price of the stock  - the closing price on the last trading day of that year for each stock'
        Dim yearly_close As Double
        
        
        'the index of the summary table         - The row of the summary table we are inputting data into'
        Dim Summary_Table_Row As Long
        
        'Yearly change                          - difference in yearly open and yearly close calculated, then stored in the summary table'
        Dim yearly_change As Double
        
        'Percent change over the year           - percentage change with respect to the opening price of that year to the closing price then stored in the summary table'
        Dim percent_change As Double
        
        
        'Name Three Columns of the summary table:'
        
        ws.Range("J" & 1).Value = "Yearly Change, $"
        ws.Range("K" & 1).Value = "Percent Change, %"
        ws.Range("L" & 1).Value = "Total Stock Volume, $"
        
        ' Format Column J as Currency (Dollars)
        ws.Columns("J:J").NumberFormat = "$#,##0.00"
    
        ' Format Column K as Percentage
        ws.Columns("K:K").NumberFormat = "0.00%"
        
        ' Format Column L as Currency (Dollars)
        ws.Columns("L:L").NumberFormat = "$#,##0.00"
        


        'Lucky for us, The data is arranged in chronological order.
        'Becase we want the information from the data located in the first row,
        'we need the script to identify when first row of data is being analysed'
        'We will use a boolean to tell us that the row is first in each stocks dataset'
        Dim first_TRUE_FALSE As Boolean
        
        
        'Create a loop that loops through the entire dataset'
        
        'Variable to identify the last row in the loop of the dataset'
        Dim LastRow As Long
        'index of current row of loop'
        Dim i As Long
        
        
        ' Initialize variables'
        TotalVolume = 0#                '- Total Volume = 0 for first trading day'
        Summary_Table_Row = 2           '- Dataset starts on row 2'
        first_TRUE_FALSE = True         '- Dataset begins with yearly open for each stock'
        
        ' Find the last row of the range in column 1
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Loop through the data
        For i = 2 To LastRow
        'The dataset contains the values for each stock in a block or rows then moves onto the next stock;
        'this means the last rows data will have a ticker symbol that does not equal the nexts rows stock.
        
        'To identify the last row of a stock within a block, the next row will not equal the current row'
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then    '-Tells us this is the last row of a particular stock'
                
                ' Set the current Ticker as the one that has been totaled so far
                Ticker = ws.Cells(i, 1).Value
                
                ' Add the final volume value to the total volume value
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                'As it is the last value in a block of chronological order, the yearl_close will equal the close of that day'
                yearly_close = ws.Range("F" & i).Value
                
                
                'Use the yearly_close, yearly_open to find the yearly change and that years percent change of stock'
                yearly_change = yearly_close - yearly_open
                percent_change = (yearly_change / yearly_open)
                
                
                ' Put values into the summary table:
                
                'Ticker value'
                ws.Cells(Summary_Table_Row, 9).Value = Ticker
                
                'Total volume'
                ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                
                'yearly_change'
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                
                'Format the yearly change column so that a negative value is shown in red and postive value is shown in green.
                'If value less than zero, then make cell color red, otherwise make cell color green.'
                If ws.Range("J" & Summary_Table_Row).Value < 0# Then
                            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                        Else
                            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                    End If
                'Percent_change'
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                
                
                ' As it was the last row of a stock, Variables must be reset for the next stock:
                'Reset the total volume for the next stock
                TotalVolume = 0
                
                'Next iteratation will be the first row for the next stock'
                first_TRUE_FALSE = True
                
                ' Change to the next row in the summary table for new stock'
                Summary_Table_Row = Summary_Table_Row + 1
            
            
            
            
            ' Determine if the current row i is equal to the next row, i.e. not the last row within a block of a stock'
            Else: ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value
                
                ' Adds the current row value of spend to Brand Total
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                If first_TRUE_FALSE = True Then
                    yearly_open = ws.Range("C" & i).Value
                    first_TRUE_FALSE = False
                    End If
            
            End If
        Next i
    Next ws
End Sub
