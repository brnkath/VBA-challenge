Attribute VB_Name = "Module11"
Sub stock_analysis()
    ' Run process across all sheats/years
    For Each ws In Worksheets
        ' Add applicable column headers to each sheet/year
        ws.Range("$I1").value = "Ticker"
        ws.Range("$J1").value = "Yearly Change"
        ws.Range("$K1").value = "Percent Change"
        ws.Range("$L1").value = "Total Stock Volume"
        ws.Range("$P1").value = "Ticker"
        ws.Range("$Q1").value = "Value"
        ws.Range("$O2").value = "Greatest % Increase"
        ws.Range("$O3").value = "Greatest % Decrease"
        ws.Range("$O4").value = "Greatest Total Volume"
        
        ' Set up variables to hold the sheet name and last row
        Dim worksheet_name As String
        worksheet_name = ws.Name
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set up variables for ticker symbol column
        Dim ticker_symbol As String
        Dim ticker_location As Integer
        ticker_location = 2
        
        ' Set up variables for ticker prices/changes
        Dim open_price As Double
        open_price = 0
        Dim open_date As Long
        Dim close_price As Double
        close_price = 0
        Dim close_date As Long
        Dim yearly_change As Double
        yearly_change = 0
        Dim percent_change As Double
        percent_change = 0
        
        ' Set up variables for total volume column
        Dim total_volume As Double
        total_volume = 0
        Dim value As Double
        
        ' Set up loop to perform stock analysis
        For i = 2 To last_row
        
            ' Establish open price for each stock
            If ws.Cells(i, 1).value <> ws.Cells(i - 1, 1).value Then
            
                open_price = ws.Cells(i, 3).value
                
            End If
        
            ' Compare each ticker symbol against the previous to see if we are still in the same ticker
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
            
                ' Set the ticker symbol
                ticker_symbol = ws.Cells(i, 1).value
                
                ' Add to the total volume
                total_volume = total_volume + ws.Cells(i, 7).value
                
                ' Add the ticker symbol to column I
                ws.Range("I" & ticker_location).value = ticker_symbol
                
                ' Add the total volume to column L
                ws.Range("L" & ticker_location).value = total_volume
                
                ' Add to the ticker location
                ticker_location = ticker_location + 1
                
                ' Reset the volume total
                total_volume = 0
                
            ' If the ticker symbol is the same as the one above
            Else
                
                ' Add to the total_volume
                total_volume = total_volume + ws.Cells(i, 7).value
                
            End If
            
            ' Establish close price for each stock and calculate/transfer change information
            If ws.Cells(i, 1).value <> ws.Cells(i + 1, 1).value Then
            
                close_price = ws.Cells(i, 6).value
                
                yearly_change = close_price - open_price
                
                ws.Range("J" & ticker_location - 1).value = yearly_change
                
                If (yearly_change <= 0) Then
                
                    ws.Range("J" & ticker_location - 1).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Range("J" & ticker_location - 1).Interior.ColorIndex = 4
                    
                End If
                
                percent_change = yearly_change / open_price
                
                ws.Range("K" & ticker_location - 1).value = percent_change
                
                If (percent_change <= 0) Then
                
                    ws.Range("K" & ticker_location - 1).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Range("K" & ticker_location - 1).Interior.ColorIndex = 4
                    
                End If
                
            End If
            
            ' Reset variables
            open_date = ws.Cells(i, 2).value
            close_date = ws.Cells(i, 2).value
            yearly_change = 0
            percent_change = 0
            
        Next i
        
        ' Set column P values
        ws.Cells(2, 16).value = Application.VLookup(ws.Range("Q2"), ws.Range("I2:L5000"), 1, False)
        ws.Cells(3, 16).value = Application.VLookup(ws.Range("Q3"), ws.Range("I2:L5000"), 1, False)
        ws.Cells(4, 16).value = Application.VLookup(ws.Range("Q4"), ws.Range("I2:L5000"), 4, False)
        
        ' Set column Q values
        ws.Cells(2, 17).value = WorksheetFunction.Max(ws.Range("$K$2:$K$5000"))
        ws.Cells(3, 17).value = WorksheetFunction.Min(ws.Range("$K$2:$K$5000"))
        ws.Cells(4, 17).value = WorksheetFunction.Max(ws.Range("$L$2:$L$5000"))
        
    ' Move on to next sheet/year
    Next ws
End Sub
