Sub doTheThing()
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"

Dim ticker As String

Dim total As Double
total = 0

Dim table_row As Integer
table_row = 2

Dim Lastrow As Long
Lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1)
        total = total + Cells(i, 7)
        Range("I" & table_row).Value = ticker
        Range("J" & table_row).Value = total
        table_row = table_row + 1
        total = 0
    Else
        total = total + Cells(i, 7)
    End If
    
Next i

End Sub