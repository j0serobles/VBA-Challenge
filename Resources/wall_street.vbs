Sub process_sheet()

' Procedure to create new columns with this format:
' +----------------------------------------------------------------------+
' | Ticker Symbol | Yearly Change | Percent Change | Total Yearly Volume |
' +----------------------------------------------------------------------+
' for each stock in the data file.
'
' Algorith:
'  For each row loop:
'    Save the opening price for the current stock.
'    Save the closing price and accumulate the volume for that day.
'    If the next ticker symbol is not equal to the current one:
'      Update the summary column with the ticker symbol, the yearly change
'      percent change, and total volume for that particular stock.
'      (The summary values for each subsequent stock go in the next row).
'      Reset the yearlychange, percent change and yearly volume to 0.
'    End If
'  End loop
'
' init vars

Dim columnNumber As Integer
Dim tickerSymbol As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim LastRow      As Long
Dim IsFirstRow   As Boolean
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume   As LongLong
Dim currentSummaryRow As Integer

columnNumber = 1
openingPrice = 0
closingPrice = 0
yearlyChange = 0
totalVolume = 0
currentSummaryRow = 2

IsFirstRow = True
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row


' Set headings for Summary columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"


' loop through each cell in column 1

For i = 2 To LastRow
' If this is the first row to process, save the opening price
  If (IsFirstRow = True) Then
    openingPrice = Cells(i, 3).Value
    IsFirstRow = False
  End If
' Accumulate the totalVolume value
    totalVolume = totalVolume + Cells(i, 7)
'  If (value of current cell) <> (value of next cell) then update the summary values
    If (Cells(i, columnNumber).Value) <> (Cells(i + 1, columnNumber)) Then
' Set the values for the summary columns
    closingPrice = Cells(i, 6).Value
    yearlyChange = closingPrice - openingPrice
    
    If (openingPrice > 0) Then
      percentChange = (yearlyChange / openingPrice)
    End If
    
    tickerSymbol = Cells(i, columnNumber).Value
    Cells(currentSummaryRow, 9).Value = tickerSymbol
    Cells(currentSummaryRow, 10).Value = yearlyChange
'Format the Percent Change column's color accordingly
    If yearlyChange >= 0 Then
       Cells(currentSummaryRow, 11).Interior.ColorIndex = 4
    Else
       Cells(currentSummaryRow, 11).Interior.ColorIndex = 3
    End If
    Cells(currentSummaryRow, 11).Value = percentChange
    Cells(currentSummaryRow, 11).NumberFormat = "0.00%"
    Cells(currentSummaryRow, 12).Value = totalVolume
'Reset the variables for the next ticker symbol
    currentSummaryRow = currentSummaryRow + 1
    closingPrice = 0
    yearlyChange = 0
    openingPrice = 0
    totalVolume = 0
    percentChange = 0
    IsFirstRow = True
    totalVolume = 0
    
  End If
Next i

findMaxIncrease
findMaxDecrease
findMaxVolume

End Sub
Sub findMaxIncrease()
  Dim maxValue As Double
  Dim LastRow As Long
  Dim i As Long
  Dim maxTicker As String
  
  maxValue = 0#
  LastRow = Range("K1").SpecialCells(xlCellTypeLastCell).Row
  For i = 2 To LastRow
    If Cells(i, 11).Value > maxValue Then
      maxValue = Cells(i, 11).Value
      maxTicker = Cells(i, 9)
    End If
  Next i
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"
  Range("O2").Value = "Greatest % Increase"
  Range("P2").Value = maxTicker
  Range("Q2").Value = maxValue
  Range("Q2").NumberFormat = "0.00%"
    
End Sub
Sub findMaxDecrease()

  Dim minValue As Double
  Dim LastRow As Long
  Dim i As Long
  Dim minTicker As String
  
  minValue = 0#
  LastRow = Range("K1").SpecialCells(xlCellTypeLastCell).Row
  For i = 2 To LastRow
    If Cells(i, 11).Value < minValue Then
      minValue = Cells(i, 11).Value
      minTicker = Cells(i, 9)
    End If
  Next i

  Range("O3").Value = "Greatest % Decrease"
  Range("P3").Value = minTicker
  Range("Q3").Value = minValue
  Range("Q3").NumberFormat = "0.00%"
  
End Sub
Sub findMaxVolume()

  Dim maxValue As Double
  Dim LastRow As Long
  Dim i As Long
  Dim maxTicker As String
  
  maxValue = 0#
  LastRow = Range("L1").SpecialCells(xlCellTypeLastCell).Row
  For i = 2 To LastRow
    If Cells(i, 12).Value > maxValue Then
      maxValue = Cells(i, 12).Value
      maxTicker = Cells(i, 9)
    End If
  Next i

  Range("O4").Value = "Greatest Total Volume"
  Range("P4").Value = maxTicker
  Range("Q4").Value = maxValue
  Range("Q4").NumberFormat = "###,###,###,###,###"
    
End Sub
Sub main()
 'traverse the list of worksheets
  Dim Current As Worksheet
  For Each Current In Worksheets
    Current.Activate
    process_sheet
  Next
End Sub

