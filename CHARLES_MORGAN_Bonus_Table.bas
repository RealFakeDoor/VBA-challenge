Attribute VB_Name = "Bonus_TABLE"
'Now we have the code for finding:
            'greatest percentage increase
            'greatest percentage decrease
            'greatest volume'
        'within the summary table:
            
        'we must create a variable to store each of the above and the corresponding ticker.
Sub Stock_Summarizer()
    'Create a variable that allows us to iterate through the worksheets'
    'WS - Current Worksheet variable'
    Dim ws As Worksheet
    
    'Loop through the worksheets with WS - Representing the current worksheet'
    For Each ws In Worksheets

        'Variable to store the greatest increase percentage and corresponding ticker'
        Dim Greatest_Increase As Double
        Dim Greatest_Increase_Ticker As String
        
        'Initiliase variable to equal the first row'
        Greatest_Increase = ws.Range("K" & 2).Value
        Greatest_Increase_Ticker = ws.Range("I" & 2).Value
        
        'Variable to store the greatest decrease percentage and corresponding ticker'
        Dim Greatest_Decrease As Double
        Dim Greatest_Decrease_Ticker As String
        
        'Initiliase variable to equal the first row'
        Greatest_Decrease = ws.Range("K" & 2).Value
        Greatest_Decrease_Ticker = ws.Range("I" & 2).Value
        
        'Variable to store the largest total volume and corresponding ticker'
        Dim Greatest_Volume As Double
        Dim Greatest_Volume_Ticker As String
        
        'Initiliase variable to equal the first row'
        Greatest_Volume = ws.Range("L" & 2).Value
        Greatest_Volume_Ticker = ws.Range("I" & 2).Value
        
        
        'Now we must create a loop that goes through the summary table and pulls
        'out the aforementioned variables and stores them'
        
        
        ' Define an index variable for iterating through the summary table,
        Dim LastRow_Summary_Table As Long
        
        'Find the last row of the range in column 1'
        LastRow_Summary_Table = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        
        'Loop through the summary table'
        For i = 2 To LastRow_Summary_Table
            'As loop iterates through find the greatest value,
            'IF value found that is greater than the value currently stored, it must then be stored instead.'
            If ws.Range("K" & i).Value > Greatest_Increase Then
                        Greatest_Increase = ws.Range("K" & i).Value
                        Greatest_Increase_Ticker = ws.Range("I" & i).Value
                End If
            
            'As loop iterates through find the lowest value,
            'IF the value found that is less than the Value currently stored, it must then be stored instead'
            If ws.Range("K" & i).Value < Greatest_Decrease Then
                        Greatest_Decrease = ws.Range("K" & i).Value
                        Greatest_Decrease_Ticker = ws.Range("I" & i).Value
                End If
            
            'As loop iterates through find the greatest value,
            'IF value found that is greater than the value currently stored, it must then be stored instead.
            If ws.Range("L" & i).Value > Greatest_Volume Then
                        Greatest_Volume = ws.Range("L" & i).Value
                        Greatest_Volume_Ticker = ws.Range("I" & i).Value
                End If
            
            Next i
        
        'Create a new table for the variables'
        'Name the columns and rows for the new table'
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("Q" & 1).Value = "Value"
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        
        
        'Put the Greatest Percentage increase value and ticker into new table,
        'format the table cell as a percentage'
        ws.Range("P" & 2).Value = Greatest_Increase_Ticker
        ws.Range("Q" & 2).Value = Greatest_Increase
        ws.Range("Q" & 2).NumberFormat = "0.00%"
        
        'Put the Greatest percentage decrease value and ticker into new table,
        'format the table cell as a percentage'
        ws.Range("P" & 3).Value = Greatest_Decrease_Ticker
        ws.Range("Q" & 3).Value = Greatest_Decrease
        ws.Range("Q" & 3).NumberFormat = "0.00%"
        
        'Put the Greatest total volume value and ticker into new table,
        'format the table cell as a currency, dollar amount'
        ws.Range("P" & 4).Value = Greatest_Volume_Ticker
        ws.Range("Q" & 4).Value = Greatest_Volume
        ws.Range("Q" & 4).NumberFormat = "$#,##0.00"
        
    Next ws
End Sub

