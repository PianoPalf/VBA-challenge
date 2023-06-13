# Stock Data Calculations/Summaries - VBA Script

Simple script for calculating and summarising Stock Data in Excel using VBA to loop through each row by Ticker.

## Table of Contents

- [General info](#general-info)
- [Technologies](#technologies)
- [Setup](#setup)
- [Screenshot](#screenshot)
- [Code example](#code-example)
- [References](#references)

## General info

- Calculates Yearly Change by Ticker: outputs data and highlights cells green for net gain or red for net loss.

- Calculates and outputs Percent Change based on Yearly Change.

- Tallies Total Stock Volume for each Ticker.

- Determines and displays Ticker with Greatest Percent Increase, Decrease and Total Stock Volume in Summary Table.

- Loops through each worksheet in Excel file.

- Created and submitted for an assignment for Monash University Data Analytics Boot Camp (June 2023).

## Technologies

Project created and run using:

- Microsoft Excel for Mac Version 16.73
- Microsoft Visual Basic for Applications 7.1

## Setup 

Use VBA Script on Excel files with Columns containing:

- 'tickers', 'date', 'open', 'low', 'high', 'close' in this order.
- See Screenshot below for more info.

## Screenshot

![screenshot](https://github.com/PianoPalf/VBA-challenge/assets/119825935/61136f19-5748-4c53-9dc3-1100a508269b)

LEFT - Input 		/ 		MIDDLE - Output (Yearly Change, Percent Change, Total Stock Volume)		/		RIGHT - Output (Summary Table)

## Code example 

```vbscript
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
```

## References

- Code snippet to Loop through worksheets adapted from: https://excelmacromastery.com/excel-vba-worksheet/
- Code, in general, was adapted from Monash University Data Analytics Boot Camp 2023 course learning material.



Created and written by Samuel Palframan - June 2023.
