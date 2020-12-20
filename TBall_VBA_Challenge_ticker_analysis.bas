Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.


'The ticker symbol.
'Done

'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'Done

'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'Almost

'The total stock volume of the stock.
'Almost



'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'Done


Sub test():
   
    For Each ws In Worksheets
    
    'Create total_vol variable as Long
    Dim total_vol As LongLong
    total_vol = 0
    ws.Cells(1, 12).Value = "Total Volume"
    
    'create dimmension for start
    Dim start_price As Double
    start_price = 2
    
    'create variable for percent_change
    Dim percent_change As Double
    percent_change = 0
    ws.Cells(1, 11).Value = "Percent Change"
       
    'set total's initial value to zero
    Dim total As Double
    total = 0

    'set variable for holding ticker symbol
    Dim ticker_symbol As String
    ws.Cells(1, 9).Value = "Ticker"
        
    'Create variable for the yearly change in price
    Dim yearly_change As Double
    yearly_change = 0
    ws.Cells(1, 10).Value = "Yearly Change"
    
    'keep track of the location for each ticker in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'create lastrow function
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    'Loop through all tickers
    For i = 2 To last_row
    
     'Check if the ticker has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set ticker symbol
            ticker_symbol = ws.Cells(i, 1).Value
            
            'Determine the yearly_change
            yearly_change = (ws.Cells(i, 6) - ws.Cells(start_price, 3))
            
            'Determine percent_change
            percent_change = Round((yearly_change / ws.Cells(start_price, 3) * 100), 2)
            
            'Determine total_vol
            total_vol = total_vol + Cells(i, 7).Value
                                                                
                  
            'Print ticker_symbol in the Summary Table
            ws.Range("I" & summary_table_row).Value = ticker_symbol
            
            'Print annual_change in the Summary Table
            ws.Range("J" & summary_table_row).Value = yearly_change
            
            'Print percent_change in the Summary Table
            ws.Range("K" & summary_table_row).Value = percent_change
            
            'Print total_vol in the Summary Table
            ws.Range("L" & summary_table_row).Value = total_vol
            
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'Reset annual_change
            yearly_change = 0
            'stock_volume = 0
            'price_difference = 0
            
        'set the starting price for the next iteration (ticker) for the cell (i+1,...)
        start_price = i + 1
        
        'if cell immediately following a row is the same brand
        Else
        
            'yearly_change = yearly_change + ws.Cells(last_row, 6).Value - ws.Cells(start_price, 3).Value
            total_vol = total_vol + ws.Cells(i, 7).Value
            
    
            
        End If
        
    Next i
    
    Next

    
End Sub
