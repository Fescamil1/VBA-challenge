Attribute VB_Name = "Module1"
Sub stocks():
    'declare vaiables
    Dim current_row, lrow, lcol, total_vol, results_row As Long
    Dim open_price, close_price As Double
    Dim gincrease, gdecrease, gvolume As Double
    Dim ginc_ticker, gdec_ticker, gvol_ticker As String
    Dim yearly_change, per_change As Double
    Dim ws As Worksheet
        
     'start loop for worksheet loop
    For Each ws In ThisWorkbook.Worksheets
            
           'Find the last non-blank cell in column A(1)
            lrow = ws.Cells(Rows.Count, 1).End(xlUp).row
            
            'Find the last non-blank cell in row 1, will be used to set the headers for the results
            lcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        
            'label the columns for the results one column after empty column
            ws.Cells(1, lcol + 1).Value = "Ticker"
            ws.Cells(1, lcol + 2).Value = "Yearly Change"
            ws.Cells(1, lcol + 3).Value = "Percent Change"
            ws.Cells(1, lcol + 4).Value = "Total Stock Volume"
            
            'add greatest change section titles
            ws.Cells(2, lcol + 6).Value = "Greatest % Increase"
            ws.Cells(3, lcol + 6).Value = "Greatest % Decrease"
            ws.Cells(4, lcol + 6).Value = "Greatest Total Volume"
            ws.Cells(1, lcol + 7).Value = "Ticker"
            ws.Cells(1, lcol + 8).Value = "Value"
            
            'set values prior to loop
            total_vol = 0
            results_row = 2
            gincrease = 0
            gdecrease = 0
            gvolume = 0
            
            'set the 1st of the opening price
            open_price = ws.Cells(2, 3).Value
            
            For current_row = 2 To lrow
                total_vol = total_vol + ws.Cells(current_row, 7).Value 'Add volume amount
                
                If ws.Cells(current_row + 1, 1).Value <> ws.Cells(current_row, 1).Value Then  'find when the ticker name changes
                         ws.Cells(results_row, lcol + 1).Value = ws.Cells(current_row, 1).Value  'add the name of the new ticker
                         ws.Cells(results_row, lcol + 4).Value = total_vol  'set total volume for current ticker
                         
                         'Get closing price, calculate the changes and set the results
                         close_price = ws.Cells(current_row, 6).Value
                         yearly_change = close_price - open_price
                         ws.Cells(results_row, lcol + 2).Value = yearly_change
                        
                         
                         'conditional formatting that will highlight positive change in green and negative change in red.
                         If yearly_change > 0 Then
                             ws.Cells(results_row, lcol + 2).Interior.ColorIndex = 4
                         Else
                             ws.Cells(results_row, lcol + 2).Interior.ColorIndex = 3
                         End If
                         
                         If yearly_change = 0 Then
                            per_change = 0
                         ElseIf open_price = 0 Then
                            per_change = (yearly_change / 0.000001) 'Address the divide by 0 issue
                         Else
                           per_change = (yearly_change / open_price)
                         End If
                         
                         ws.Cells(results_row, lcol + 3).NumberFormat = "0.00%" 'set format to show as % in output
                         ws.Cells(results_row, lcol + 3).Value = per_change
                         
                         'look for greatest increase
                         If per_change > gincrease Then
                             gincrease = per_change
                             ginc_ticker = ws.Cells(current_row, 1).Value
                         End If
                         
                         'look for greatest decrease
                         If per_change < gdecrease Then
                             gdecrease = per_change
                            gdec_ticker = ws.Cells(current_row, 1).Value
                         End If
                         'look for greatest volume
                         If total_vol > gvolume Then
                             gvolume = total_vol
                             gvol_ticker = ws.Cells(current_row, 1).Value
                         End If
                                
                         'reset variables and totals
                         results_row = results_row + 1   'set the pointer for the next results row
                         total_vol = 0                   'reset the volume to 0
                         open_price = ws.Cells(current_row + 1, 3).Value 'reset open_price for next
                End If
            Next current_row
            
            'set the values for greatest values found
            ws.Cells(2, lcol + 7).Value = ginc_ticker
            ws.Cells(2, lcol + 8).Value = gincrease
            ws.Cells(2, lcol + 8).NumberFormat = "0.00%"
            ws.Cells(3, lcol + 7).Value = gdec_ticker
            ws.Cells(3, lcol + 8).Value = gdecrease
            ws.Cells(3, lcol + 8).NumberFormat = "0.00%"
            ws.Cells(4, lcol + 7).Value = gvol_ticker
            ws.Cells(4, lcol + 8).Value = gvolume
    
    Next ws
    
End Sub
