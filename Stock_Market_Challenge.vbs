Sub StockMarket():

'To prevent overflow error on stock volume variable
On Error Resume Next

'Execute script on each worksheet at once
For Each ws In Worksheets
    
    'Declare variables for summary table
    Dim ticker As String
    
    Dim EndRow As Long
    
    Dim vol As Long
    
    'Start with volume = 0
    vol = 0
    
    Dim year_open As Double
    
    Dim year_close As Double
    
    Dim percent_change As Double
    
    Dim summary_table_row As Long
    
    'Start at row 2 & increment by 1 after each loop
    summary_table_row = 2
    
    'Declare variables for second summary table
    Dim greatest_volume As Long
    
    Dim greatest_increase As Double
    
    Dim greatest_decrease As Double
    
    Dim EndRow9 As Long
    
    Dim ws_name As String
    
    'Create headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Determine the last non-blank row in column A
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through entire column A
    For i = 2 To EndRow
        
        'If 2 consecutive ticker values are different, then...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Get ticker & calculate yearly change
            ticker = ws.Cells(i, 1).Value
            
            year_open = ws.Cells(summary_table_row, 3).Value
            
            year_close = ws.Cells(i, 6).Value
            
            'Calculate total volume for given ticker
            vol = vol + ws.Cells(i, 7).Value
            
            'Write ticker & yearly change values to summary table
            ws.Cells(summary_table_row, 9).Value = ticker
            
            ws.Cells(summary_table_row, 10).Value = year_close - year_open
                
                'Set conditional formatting for yearly change column
                If ws.Cells(summary_table_row, 10).Value > 0 Then
                
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    
                Else
                
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Set conditional formatting for percent change column
                If ws.Cells(summary_table_row, 3).Value <> 0 Then
                
                    percent_change = (year_close - year_open) / (year_open)
                    
                    ws.Cells(summary_table_row, 11).Value = Format(percent_change, "Percent")
                    
                Else
                
                    ws.Cells(summary_table_row, 11).Value = Format(0, "Percent")
                
                End If
                
            'Write total volume to summary table, reset vol to 0 & increment to next row in summary table
            ws.Cells(summary_table_row, 12).Value = vol
        
            summary_table_row = summary_table_row + 1
        
            vol = 0
        
        Else
            
            'Continue adding to total volume
            vol = vol + ws.Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Find last non-blank cell in column 9
    EndRow9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Set desired outputs equal to values in first row of the new loop, then continuously check for greater values
    greatest_vol = ws.Cells(2, 12).Value
    
    greatest_increase = ws.Cells(2, 11).Value
    
    greatest_decrease = ws.Cells(2, 11).Value
    
    'Start loop through column 9
    For i = 2 To EndRow9
       
       'Update outputs if next i is greater than previous i
        If ws.Cells(i, 12).Value > greatest_vol Then
       
        greatest_vol = ws.Cells(i, 12).Value
       
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
       
        'If next i is <= previous i, do nothing
       Else
       
        greatest_vol = greatest_vol
       
       End If
       
       'Create similar If statements for 2 remaining outputs
       If ws.Cells(i, 11).Value > greatest_increase Then
       
        greatest_increase = ws.Cells(i, 11).Value
       
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    
       Else
        
        greatest_increase = greatest_increase
        
       End If
       
       If ws.Cells(i, 11).Value < greatest_decrease Then
       
        greatest_decrease = ws.Cells(i, 11).Value
       
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    
       Else
        
        greatest_decrease = greatest_decrease
        
       End If
         
        'Format values in summary table
        ws.Cells(4, 17).Value = Format(greatest_vol, "Scientific")
        ws.Cells(2, 17).Value = Format(greatest_increase, "Percent")
        ws.Cells(3, 17).Value = Format(greatest_decrease, "Percent")
        
    Next i
    
    Worksheets(ws_name).Columns("A:Z").AutoFit
    
Next ws

End Sub