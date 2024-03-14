'Part 1: Ticker Column

'Copy Column A contents into Column I
Sub TickerColumn()
    Range("A2:A22771").Copy Range("i2:i22771")
End Sub

'Part 2: Yearly Change

Sub YearlyChange()
Dim i As Long
Dim Total As Double

For i = 2 To 753001


'Closing Price - Open Price and populate answer in Yearly Change column
    Total = Cells(i, 6).Value - Cells(i, 3).Value
    
    Cells(i, 10) = Total

    
    'Nest color fill for each cell: Red(3) for negative and Green(4) for positive
    If Cells(i, 10).Value >= 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
    Else: Cells(i, 10).Interior.ColorIndex = 3
    End If
    
 'Loop through each row
 Next i

End Sub

'Part 3: Percent Change

Sub Percent_Change()
Dim i As Long
Dim Total As Double

For i = 2 To 1500

    '(Closing Price - Open Price)/Opening Price *100 & populate answer in Percent Change column
    Total = (Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value
    
    Cells(i, 11) = Total
    Cells(i, 11).NumberFormat = "0.00%"
    
Next i
    
End Sub

'Part 4: Total Stock Volume

Sub Total_Stock_Volume()

Dim i As Long
Dim Total As Double

For i = 2 To 753001

    'Close Price * Volume and populate answer in Total Stock Volume Column
    Total = Cells(i, 6).Value * Cells(i, 7).Value
    
    Cells(i, 12) = Total
    
Next i

End Sub





'______________________________________
'Unsuccessful attempts
'______________________________________

'part 6 (1)
Sub TestMax()


Dim ws As Worksheet
Dim currentCell As Range
Dim maxval As Double
Dim lastRow As Long

    'Set worksheet to find highest value
    Set ws = ActiveSheet.a

 
 'Set maxval
    ws.Range("Q2").Value = maxval

    
    'Set last row in column
    lastRow = Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    
  ' Loop through rows in the column
    For Each currentCell In ws.Range("K2:K" & lastRow)

    ' Searches for when the value of the next cell is different than that of the current cell
        If IsNumeric(currentCell.Value) Then
            If currentCell.Value > maxval Then
                maxval = currentCell.Value
            End If
        End If
    Next currentCell
    
    'Display the highest value found
    maxval = ws.Range("Q2").Value
    ws.Range("Q2").NumberFormat = "0.00%"
      
  
End Sub

'Part 6 (2)
Sub GreatestTotalVol()

Dim high As Double
Dim rng As Range

rng = alphabetical_testing.Sheets("A").Range("K2:K22771")

high = Application.WorksheetFunction.Max(rng)

ThisWorkbook.Sheets("A").Range("Q2").Value = high

End Sub




'Part 6Take (3)
Sub MaxTest()
    Dim ws As Worksheet
    Dim currentCell As Range
    Dim maxval As Double
    Dim lastRow As Long
    
    ' Set worksheet to find highest value. Replace "ThisWorkbook" with your workbook variable if necessary.
    Set ws = alphabetical_testing.Sheets("A")
    
    ' Initialize maxval
    maxval = -99999
    
    ' Set last row in column K
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Loop through rows in column K to find the max value
    For Each currentCell In ws.Range("K2:K" & lastRow)
        ' Checks if the cell contains a numeric value and if it's greater than maxval
        If IsNumeric(currentCell.Value) Then
            If currentCell.Value > maxval Then
                maxval = currentCell.Value
            End If
        End If
    Next currentCell
    
    ' Display the highest value found in cell Q2
    ws.Range("Q2").Value = maxval
    ' Format as percentage (optional, adjust as needed)
    ws.Range("Q2").NumberFormat = "0.00%"
End Sub