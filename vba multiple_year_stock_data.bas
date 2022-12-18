Attribute VB_Name = "Module3"
' 1.)Script that loops through all stocks for one year and outputs following information on all worksheets:
    ' The ticker symbol
    ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.
' 2.)BONUS: Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"

'Begin Sub procedure
Sub Multiple_year_stock_data()

    'Activate loop for all worksheets
    For Each ws In Worksheets
        
        'Declare Dimensions/Variables
        Dim tickername As String
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim tickervolume As Double
            tickervolume = 0
        Dim summary_ticker_row As Integer
            summary_ticker_row = 2
            open_price = Cells(2, 3).Value
        Dim close_price As Double

        'Label new column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
            
        'Find last row of ticker column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

        'For loop through year by ticker name
        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              tickername = ws.Cells(i, 1).Value

              ws.Range("I" & summary_ticker_row).Value = tickername

              ws.Range("L" & summary_ticker_row).Value = tickervolume

              tickervolume = tickervolume + ws.Cells(i, 7).Value

              close_price = ws.Cells(i, 6).Value

              yearly_change = (close_price - open_price)
              
              ws.Range("J" & summary_ticker_row).Value = yearly_change
              
              'Yearly Change
                If (open_price = 0) Then
                    percent_change = 0
                    
                Else
                    percent_change = yearly_change / open_price
                
                End If

              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              summary_ticker_row = summary_ticker_row + 1
              
              tickervolume = 0

              open_price = ws.Cells(i + 1, 3)
            
            Else
              
              tickervolume = tickervolume + ws.Cells(i, 7).Value
            
            End If
        
        Next i

              lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    'Yearly Change color code
    For i = 2 To lastrow_summary_table
        If ws.Cells(i, 10).Value > 0 Then
           ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
           ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'maximum/minimum percent change, max volume
        greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
        greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
        greatest_volume = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        ws.Range("Q2") = "%" & greatest_increase * 100
        ws.Range("Q3") = "%" & greatest_decrease * 100
        ws.Range("Q4") = greatest_volume

        'find ticker symbol matching values above
        inc_loc = WorksheetFunction.Match(greatest_increase, ws.Range("K2:K" & RowCount), 0)
        dec_loc = WorksheetFunction.Match(greatest_decrease, ws.Range("K2:K" & RowCount), 0)
        vol_loc = WorksheetFunction.Match(greatest_volume, ws.Range("L2:L" & RowCount), 0)

        'Assigned to cells
        ws.Range("P2") = ws.Cells(inc_loc + 1, 9)
        ws.Range("P3") = ws.Cells(dec_loc + 1, 9)
        ws.Range("P4") = ws.Cells(vol_loc + 1, 9)
        
    Next ws

End Sub
