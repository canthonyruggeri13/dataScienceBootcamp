Sub VbaHw()


Dim ticker As String
Dim vol As Double
vol = 0
Dim SummaryTableRow As Integer
SummaryTableRow = 2

Cells(1, 9).Value = "Ticker Name"
Cells(1, 10).Value = "Total Volume"



For i = 2 To 797711
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       ticker = Cells(i, 1).Value

   Range("I" & SummaryTableRow).Value = ticker
       vol = vol + Cells(i, 7).Value

    Range("J" & SummaryTableRow).Value = vol



       SummaryTableRow = SummaryTableRow + 1
      vol = 0


   Else
   vol = vol + Cells(i, 7).Value

   End If
   Next i
   
   End Sub
   


