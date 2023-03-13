Attribute VB_Name = "Module1"
' This script loops through all worksheets, and for each worksheet parses all the stocks for one year and outputs
' the following information:
' 1. The ticker symbol
' 2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 4. The total stock volume of the stock. The result should match the following image:
' 5. The stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
' 6. Highlights positive change in green and negative change in red using conditional formatting.

Sub StockDataInfoCalc()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' ------------------------------------------------------------------------------

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Variable for the yearly change per ticker
        Dim YearlyChange As Double
        YearlyChange = 0

        ' Variable for the current sum of the volume per ticker
        Dim CurrentSum As Double
        CurrentSum = 0
        
        ' Variable for the row number of opening price for each ticker
        Dim OpenRow As Long
        OpenRow = 2
        
        ' Variable for total amount of ticker of one type
        Dim TickerCount As Integer
        TickerCount = 1

        ' Variable for the greatest percent increase
        Dim MaxPerc As Double
        MaxPerc = ws.Cells(2, 11).Value
        
        ' Variable for the greatest percent decrease
        Dim MinPerc As Double
        MinPerc = ws.Cells(2, 11).Value
        
        ' Variable for the greatest total volume
        Dim MaxTotVol As Double
        MaxTotVol = ws.Cells(2, 12)
   
        ' Variable for the location of each row in the summary table
        Dim NewTableRow As Integer
        NewTableRow = 2
        
        ' ------------------------------------------------------------------------------
        
        ' Print the column headers of the new large summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Print the column headers of the new small summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ' Print the row headers of the new small summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
       
        ' Loop through all stocks in a worksheet
        For i = 2 To LastRow

            ' Check if we jump to the next ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                TickerCount = 1
                
                ' Add ticker to the new table
                 ws.Range("I" & NewTableRow).Value = ws.Cells(i, 1).Value
                             
                ' Color the cell green if the yearly change is positive, red if negative, no color if no change
                ws.Range("J" & NewTableRow).Value = ws.Cells(i, 6).Value - ws.Cells(OpenRow, 3).Value
                If (ws.Range("J" & NewTableRow).Value > 0) Then
                    ws.Range("J" & NewTableRow).Interior.ColorIndex = 4
                ElseIf (ws.Range("J" & NewTableRow).Value < 0) Then
                    ws.Range("J" & NewTableRow).Interior.ColorIndex = 3
                End If
                             
                ' Add (with "percent" format) the yearly change for the (NewTableRow)th ticker of the summary table
                ws.Range("K" & NewTableRow).Value = ((ws.Cells(i, 6).Value - ws.Cells(OpenRow, 3).Value) / (ws.Cells(OpenRow, 3).Value))
                ws.Range("K" & NewTableRow) = Format(ws.Range("K" & NewTableRow), "Percent")

                ' Keeping track of max/min values of percentage change, name of ticker
                If ((ws.Range("K" & NewTableRow).Value) >= MaxPerc) Then
                    MaxTic = ws.Cells(NewTableRow, 9).Value
                    MaxPerc = ws.Range("K" & NewTableRow).Value
                ElseIf ((ws.Range("K" & NewTableRow).Value) <= MinPerc) Then
                    MinTic = ws.Cells(NewTableRow, 9).Value
                    MinPerc = ws.Range("K" & NewTableRow).Value
                End If
                
                ws.Range("L" & NewTableRow).Value = CurrentSum
                
                ' Keep track of max value of total stock volumes, name of ticker
                If ((ws.Range("L" & NewTableRow).Value) >= MaxTotVol) Then
                    MaxTotVolTic = ws.Cells(NewTableRow, 9).Value
                    MaxTotVol = ws.Range("L" & NewTableRow).Value
                End If
                            
                NewTableRow = NewTableRow + 1
      
                CurrentSum = 0
                
                OpenRow = i + 1
            Else
                ' Count total value of stock volumes per ticker
                CurrentSum = CurrentSum + ws.Cells(i, 7).Value
                ' Count the number of tickers per ticker
                TickerCount = TickerCount + 1
            End If
        Next i
        ' ------------------------------------------------------------------------------
        ' Print the values of small summary table
        ws.Cells(2, 16).Value = MaxTic
        ws.Cells(2, 17).Value = Format(MaxPerc, "Percent")
        ws.Cells(3, 16).Value = MinTic
        ws.Cells(3, 17).Value = Format(MinPerc, "Percent")
        ws.Cells(4, 16).Value = MaxTotVolTic
        ws.Cells(4, 17).Value = MaxTotVol
                   
    Next ws
    
    MsgBox ("Summary tables generated successfully! :)")
    
End Sub


