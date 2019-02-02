Sub FirstandSecondLoop()

Dim i, n As Double

Dim ws As Worksheet

Dim TotVal As LongLong

Dim SumRow As Integer


Dim Ticker As String

For Each ws In Worksheets

SumRow = 2
TotVal = 0

n = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

    For i = 2 To n
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker = ws.Cells(i, 1).Value
        
        TotVal = TotVal + ws.Cells(i, 7)
        
        ws.Range("I" & SumRow).Value = Ticker
        ws.Range("J" & SumRow).Value = TotVal
        
        SumRow = SumRow + 1
        
        TotVal = 0
        
        Else
        
        TotVal = TotVal + ws.Cells(i, 7).Value
        
        End If
        
        

    Next i
    

Next ws



End Sub
••••ˇˇˇˇ