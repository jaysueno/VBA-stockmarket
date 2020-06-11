Sub market_breakdown()

For Each ws In Worksheets

'name new columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("k:k").NumberFormat = "0.00%"
        ws.Range("L1").Value = "Total Stock Volume"

'set an initial variable for holding ticker name
 Dim ticker As String

'set an initial variable for holding stock value change
 Dim stock_value As Double
 stock_value = 0

 Dim stock_open As Double
 stock_open = 0

 Dim stock_close As Double
 stock_close = 0
 
 Dim Vol As Double
 Vol = 0

'keep track of the location for each ticker name in the summary table
 Dim summary_table_row As Double
 summary_table_row = 2

'For LOOP with Last row function
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      For i = 2 To lastrow
        
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          stock_open = ws.Cells(i, 3)
            
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          ticker = ws.Cells(i, 1).Value

          stock_close = ws.Cells(i, 6).Value
          
          stock_value = stock_value + ws.Cells(i, 3).Value
          
          yearly_change = stock_close - stock_open

            If stock_open = 0 Then
            
                percent_change = "null"
            
                Else
            
                percent_change = (stock_close - stock_open) / stock_open
          
            End If
          
          Vol = Vol + ws.Cells(i, 7).Value

          ws.Range("I" & summary_table_row).Value = ticker

          ws.Range("J" & summary_table_row).Value = yearly_change
            If ws.Range("j" & summary_table_row).Value > 0 Then
                ws.Range("j" & summary_table_row).Interior.ColorIndex = 4
            Else
                ws.Range("j" & summary_table_row).Interior.ColorIndex = 3
            End If

          ws.Range("k" & summary_table_row).Value = percent_change
          
          ws.Range("l" & summary_table_row).Value = Vol
          
          summary_table_row = summary_table_row + 1

          Vol = 0
          stock_open = 0
          stock_close = 0
          

        Else
          Vol = Vol + ws.Cells(i, 7).Value

        End If

      Next i
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"

  lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
     
    
    Dim gpi As Double
    Dim gpd As Double
    Dim gtv As Double
    
    Dim gpi_tic As String
    Dim gpd_tic As String
    Dim gtv_tic As String
    Set my_range = Range("i:l")
    
        gpi = Application.WorksheetFunction.Max(Range("k:k"))
        ws.Range("Q2").Value = gpi
        ws.Range("Q2").NumberFormat = "0.00%"
        
        gpd = Application.WorksheetFunction.Min(Range("k:k"))
        ws.Range("Q3").Value = gpd
        ws.Range("Q3").NumberFormat = "0.00%"
        
        gtv = Application.WorksheetFunction.Max(Range("l:l"))
        ws.Range("Q4").Value = gtv
            
        'gpd_tic = Application.WorksheetFunction.VLookup(gpd, my_range, 2).Value
        'MsgBox (gpd_tic)



Next ws


End Sub

