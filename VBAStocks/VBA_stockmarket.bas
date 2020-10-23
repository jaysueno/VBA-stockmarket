Attribute VB_Name = "Module1"
'Written by Jay Sueno
'Please visit my LinkedIn - https://www.linkedin.com/in/jay-sueno-359a274/

'Create function
Sub market_breakdown()

'Start for loop through all the worksheets
For Each ws In Worksheets

  'Define new column names for the output values
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("k1").Value = "Percent Change"
  ws.Range("k:k").NumberFormat = "0.00%"
  ws.Range("L1").Value = "Total Stock Volume"

  'Define variables and their variable types
  Dim ticker As String
  Dim stock_value As Double
  Dim stock_open As Double
  Dim stock_close As Double
  Dim Vol As Double

  'Set starting values of variables
  stock_value = 0
  stock_open = 0
  stock_close = 0
  Vol = 0

  'Keep track of the location for each ticker name with this variable
  Dim summary_table_row As Double
  'Set starting value to be the second row to avoid headers
  summary_table_row = 2

  'Define the "lastrow" with the "xlup" function
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  'Begin for loop and use lastrow
  For i = 2 To lastrow
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then 'Detect if values are different
      stock_open = ws.Cells(i, 3) 'Log opening value
    End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'Detect if ticker names are different
      ticker = ws.Cells(i, 1).Value 'Assign ticker name to variable
      stock_close = ws.Cells(i, 6).Value 'Log stock closing value
      stock_value = stock_value + ws.Cells(i, 3).Value 'Calculate stock value with aggregator function
      yearly_change = stock_close - stock_open 'Calculate yearly change 

        'Function to deal with absent denominator error
        If stock_open = 0 Then
            percent_change = "null"
            Else
            percent_change = (stock_close - stock_open) / stock_open
        End If
      
      'Aggregator function
      Vol = Vol + ws.Cells(i, 7).Value
      ws.Range("I" & summary_table_row).Value = ticker
      ws.Range("J" & summary_table_row).Value = yearly_change
        If ws.Range("j" & summary_table_row).Value > 0 Then
            ws.Range("j" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("j" & summary_table_row).Interior.ColorIndex = 3
        End If

      'Populate values in output columns
      ws.Range("k" & summary_table_row).Value = percent_change
      ws.Range("l" & summary_table_row).Value = Vol

      'Increment counter
      summary_table_row = summary_table_row + 1

      'Reset values to 0
      Vol = 0
      stock_open = 0
      stock_close = 0
      
    Else
      Vol = Vol + ws.Cells(i, 7).Value
    End If
  'Go to new row
  Next i

  'Define summary table header rows and column names  
  ws.Range("p1").Value = "Ticker"
  ws.Range("q1").Value = "Value"
  ws.Range("o2").Value = "Greatest % Increase"
  ws.Range("o3").Value = "Greatest % Decrease"
  ws.Range("o4").Value = "Greatest Total Volume"

  'Defind summary table variables  
  Dim gpi As Double
  Dim gpd As Double
  Dim gtv As Double
  'Assign values to variables by finding max or min from rows "k" and "l"
  gpi = Application.WorksheetFunction.Max(Range("k:k"))
  gpd = Application.WorksheetFunction.Min(Range("k:k"))
  gtv = Application.WorksheetFunction.Max(Range("l:l"))
  
  'Begin for loop to populate summary table
  For i = 2 To summary_table_row
    If (ws.Cells(i, 11).Value = gpi) Then 'if the cell equals the desired value then populate the summary table
      ws.Range("P2").Value = ws.Cells(i, 9).Value
      ws.Range("Q2").Value = ws.Cells(i, 11).Value
      ws.Range("Q2").Style = "Percent"
    ElseIf (ws.Cells(i, 11).Value = gpd) Then 'if the cell equals the desired value then populate the summary table
      ws.Range("P3").Value = ws.Cells(i, 9).Value
      ws.Range("Q3").Value = ws.Cells(i, 11).Value
      ws.Range("Q3").Style = "Percent"
    ElseIf (ws.Cells(i, 12).Value = gtv) Then 'if the cell equals the desired value then populate the summary table
      ws.Range("P4").Value = ws.Cells(i, 9).Value
      ws.Range("Q4").Value = ws.Cells(i, 12).Value
    End If
  'Go to next row
  Next i

'Go to next worksheet      
Next ws

End Sub