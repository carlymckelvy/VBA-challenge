Attribute VB_Name = "Module1"
Sub stock_market()

    Dim ws As Worksheet
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    For Each ws In Worksheets

    'Set summary table column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Set an initial variable for holding the ticker symbol
    Dim ticker_symbol As String
    
    'Set an initial variable for holding the total per ticker
    Dim ticker_total As Variant
    ticker_total = 0
    
    'Keep track of the location for each ticker symbol in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'Determine last row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Determine first open value
    open_price = ws.Cells(2, 3).Value
    
    'Loop through all ticker lines
    For i = 2 To last_row
    
        'Check if we are still within the same ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set ticker symbol
            ticker_symbol = ws.Cells(i, 1).Value
            
            'Add to the ticker total
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
            'Print the ticker symbol in the summary table
            ws.Range("I" & summary_table_row).Value = ticker_symbol
            
            'Print ticker total in the summary table
            ws.Range("L" & summary_table_row).Value = ticker_total
          
            'Calculate opening price less closing price
            'find column c first line value
            'open_price = ws.Cells(i, 3).Value
            
            'find column f last line value
            close_price = ws.Cells(i, 6).Value
        
            yearly_change = close_price - open_price
            
            'Print change in value in column J
            ws.Range("J" & summary_table_row).Value = yearly_change
                          
            'Print percent change in column K
            If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                   
                
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            End If
                           
             'Color code red for negative and green for positive
            
                If ws.Range("J" & summary_table_row).Value >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf ws.Range("J" & summary_table_row).Value < 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
           
            
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'Reset the ticker total
            ticker_total = 0
            
            'Reset open_price
            open_price = ws.Cells(i + 1, 3).Value
            
            
            
            
            
        'If the cell immediately following a row is the same brand...
        Else
        
            'Add to the ticker total
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
    Next ws
            
            


End Sub
