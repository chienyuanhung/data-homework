Sub stockDataAnalysis()
   'declare variables
   Dim limit As Long 'number of rows, main table
   Dim i As Long
   Dim k As Long
   Dim tsv As Variant 'total stock volume
   Dim currentTicker As String
   Dim nextTicker As String
   Dim tickerName As String
   Dim xrow As Integer 'current rows
   Dim iprice As Double 'initial price
   Dim eprice As Double 'end price

   'declare variable for second table

  Dim secLimit As Integer 'number of rows, data table
  Dim xrow2 As Integer  'row number index for greatest % increase
  Dim xrow3 As Integer  'row number index for greatest % decrease
  Dim xrow4 As Integer  'row number index for greatest total volume
  Dim gpi As Double 'variable hold greatest % increase
  Dim gpd As Double 'variable hold greatest % decrease
  Dim gtv As Variant 'variable hold greatest total volumn

For each ws in worksheets
   'headers for tables
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percentage Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
   'header for second table (about the greatest change)
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"

   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"

   'initial value
   xrow = 2
   iprice = ws.Cells(2, 3).Value
   limit = ws.Cells(Rows.Count, 1).End(xlUp).Row 'number of rows, main table

   'loop throuh the data for table of total stock volume, yearly chaneg and percentage change 
   For i = 2 To limit
       currentTicker = ws.Cells(i, 1).Value
       nextTicker = ws.Cells(i + 1, 1).Value
       If currentTicker <> nextTicker Then
          'the total stock volume
          tsv = tsv + ws.Cells(i, 7)
          ws.Cells(xrow, 12).Value = tsv
          tsv = 0
          'ticker name
          tickerName = currentTicker
          currentTicker = nextTicker
          ws.Cells(xrow, 9).Value = tickerName

          'yearly Change
          eprice = ws.Cells(i, 6)
          ws.Cells(xrow, 10) = eprice - iprice
          ws.Cells(xrow, 10) = Round(ws.Cells(xrow, 10), 8)

          'percentage Change
          ws.Cells(xrow, 11) = (eprice - iprice) / iprice
          ws.Cells(xrow, 11).NumberFormat = "0.00%"
          'set new values
          xrow = xrow + 1
          if ws.Cells(i + 1, 3).value = 0 Then
             for k = (i +1) to limit  
                if ws.cells(k, 3) <> 0 then exit for
             next k 
             iprice = ws.Cells(k, 3).value
          else     
            iprice = ws.Cells(i + 1, 3)
          end if   
          
       Else
         tsv = tsv + ws.Cells(i, 7)
       End If
    Next i
    
 
  'set initial value for variable in second table
  secLimit = ws.Cells(Rows.Count, 9).End(xlUp).Row 'row number for the second table
  gpi = ws.Cells(2, 11).Value
  gpd = ws.Cells(2, 11).Value
  gtv = ws.Cells(2, 12).Value

  'change cell color (yearly change column), green for positive, red for negative
  For i = 2 To secLimit
      If ws.Cells(i, 10) >= 0 Then
         ws.Cells(i, 10).Interior.ColorIndex = 4
      Else
         ws.Cells(i, 10).Interior.ColorIndex = 3
      End If
   Next i
 
   'looping for greatest % increase
   For i = 2 To secLimit
      If ws.Cells(i, 11).Value > gpi Then
         gpi = ws.Cells(i, 11).Value
         xrow2 = i
      End If
   Next i
    
   'assign value for greatest % increase
      ws.Cells(2, 16).Value = ws.Cells(xrow2, 9).Value 'ticker for greatest % increase
      ws.Cells(2, 17).Value = ws.Cells(xrow2, 11).Value 'value for greatest % increase
      ws.Cells(2, 17).NumberFormat = "0.00%"
    
    
     'looping for greatest % decrease
   For i = 2 To secLimit
      If ws.Cells(i, 11).Value < gpd Then
         gpd = ws.Cells(i, 11).Value
         xrow3 = i
      End If
   Next i
    
    'assiagn value for greatest % decrease
    ws.Cells(3, 16).Value = ws.Cells(xrow3, 9).Value 'ticker for greatest % decrease
    ws.Cells(3, 17).Value = ws.Cells(xrow3, 11).Value 'value for greatest % decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
      'looping for greatest total volume
      
   For i = 2 To secLimit
      If ws.Cells(i, 12).Value > gtv Then
         gtv = ws.Cells(i, 12).Value
         xrow4 = i
      End If
   Next i
    
      ws.Cells(4, 16).Value = ws.Cells(xrow4, 9).Value 'ticker for greatest total volume
      ws.Cells(4, 17).Value = ws.Cells(xrow4, 12).Value 'value for greatest total volume
      ws.Cells(4, 17).NumberFormat = "General"

next ws
End Sub


