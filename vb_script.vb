Sub stock_data_calculations()

'Variable declarations
Dim ws As Worksheet
Dim yearly_change_open, yearly_change_close, yearly_change As Double
Dim ticker As String
Dim counter As Long

'Goes through each worksheet in Excel file
For Each ws In ThisWorkbook.Worksheets
    
    'Chooses active worksheet
    ws.Activate
    
    'Adds appropriate column names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Setting counters and yearly_change values
    counter = 0
    summary_table_row = 2
    yearly_change_open_counter = 2
    total_stock_volume = 0
    ticker_matches = False
    
    'Loops through each row
    For r = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
        'If Stock is different...
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            
            'Displays Ticker
            ticker = Cells(r, 1).Value
            Range("I" & summary_table_row).Value = ticker
            
            'Records Yearly Close value
            yearly_change_close = Cells(r, 6).Value
            
            'Calculates Yearly Change
            yearly_change = yearly_change_close - yearly_change_open
            Range("J" & summary_table_row).Value = yearly_change
                'Colours cell interior depending on positive or negative value
                If Range("J" & summary_table_row).Value < 0 Then
                    Range("J" & summary_table_row).Interior.Color = vbRed
                Else
                    Range("J" & summary_table_row).Interior.Color = vbGreen
                End If
            
            'Calculates and Outputs Percentage Change
            percent_change = yearly_change / yearly_change_open
            Range("K" & summary_table_row).Value = FormatPercent(percent_change)
                
        
            'Outputs Total Stock Volume
            Range("L" & summary_table_row).Value = total_stock_volume
            
            'Counters and Iterative Resets
            summary_table_row = summary_table_row + 1
            counter = 0
            total_stock_volume = 0
            ticker_matches = False
        
        'If Stock is the same...
        Else
            
            'Records Yearly Open value
            If ticker_matches = False Then
                yearly_change_open = Cells(r, 3).Value
                ticker_matches = True
            End If
            
            'Tallies Total Stock Volume
            total_stock_volume = total_stock_volume + Cells(r, 7)
                    
        End If
        
    Next r
    
    '-=== BONUS: Determining Greatest Percent Increase, Decrease and Total Stock Volume ===-
    
    'Adds appropriate column names
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Sets variables to 0
    greatest_percent = 0
    lowest_percent = 0
    greatest_volume = 0
    
    'Loops through each row of Summary Table
    For i = 2 To 91
    
        'Determines Greatest Percent Increase
        If Cells(i, 11).Value > greatest_percent Then
            
            greatest_percent = Cells(i, 11).Value
            greatest_percent_ticker = Cells(i, 9).Value
            
        'Determines Greatest Percent Decrease
        ElseIf Cells(i, 11).Value < lowest_percent Then
        
            lowest_percent = Cells(i, 11).Value
            lowest_percent_ticker = Cells(i, 9).Value
          
        'Determines Greatest Total Stock Volume
        ElseIf Cells(i, 12).Value > greatest_volume Then
            
            greatest_volume = Cells(i, 12).Value
            greatest_volume_ticker = Cells(i, 9).Value
           
        End If
        
    Next i
    
    'Outputs data to appropriate cells
    Cells(2, 16).Value = greatest_percent_ticker
    Cells(2, 17).Value = FormatPercent(greatest_percent)
    Cells(3, 16).Value = lowest_percent_ticker
    Cells(3, 17).Value = FormatPercent(lowest_percent)
    Cells(4, 16).Value = greatest_volume_ticker
    Cells(4, 17).Value = greatest_volume
    
Next ws

End Sub