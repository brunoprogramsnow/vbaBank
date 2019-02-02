Sub FirstandSecondLoop()

Dim i, n, f, j As Long
Dim ws As Worksheet

n = Cells(Rows.Count, "A").End(xlUp).Row
f = Cells(Rows.Count, "G").End(xlUp).Row

For Each ws In ThisWorkbook.Sheets

    For i = 2 To n

    ThisWorkbook.Sheets("2014").Cells(i, 9) = ThisWorkbook.Sheets("2014").Cells(i, 1).Value

    Next i
    
    For j = 2 To f
    
    ThisWorkbook.Sheets("2014").Cells(j, 10) = ThisWorkbook.Sheets("2014").Cells(j, 1).Value
    
    
    Next j
    
      For i = 2 To n

    ThisWorkbook.Sheets("2014").Cells(i, 9) = ThisWorkbook.Sheets("2014").Cells(i, 1).Value

    Next i
    
    For j = 2 To f
    
    ThisWorkbook.Sheets("2014").Cells(j, 10) = ThisWorkbook.Sheets("2014").Cells(j, 1).Value
    
    
    Next j
    
      For i = 2 To n

    ThisWorkbook.Sheets("2014").Cells(i, 9) = ThisWorkbook.Sheets("2014").Cells(i, 1).Value

    Next i
    
    For j = 2 To f
    
    ThisWorkbook.Sheets("2014").Cells(j, 10) = ThisWorkbook.Sheets("2014").Cells(j, 1).Value
    
    
    Next j

Next



End Sub