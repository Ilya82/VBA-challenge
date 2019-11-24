Sub credit_card()

  ' Set an initial variable for holding the brand name
  Dim Ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For i = 2 To 70926

    ' Check if we are still within the same Ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set Ticker
      Ticker = Cells(i, 1).Value

      ' Add Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print Total Stock Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Brand Total
      Brand_Total = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
   
    End If

  Next i

End Sub
