# VBA-challenge

Thomas Babjak

13Dec2020

(See Attachments)

________________________

Sub WorksheetLoop()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"

    Dim lastrow As Long
    Dim RowCounter As LongLong

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    RowCounter = 2

    Dim OpenStock As Double
    Dim CloseStock As Double

    OpenStock = ws.Cells(2, 3).Value

    Dim TotalVolume As LongLong

    TotalVolume = 0

    For I = 2 To lastrow

        TotalVolume = ws.Cells(I, 7).Value + TotalVolume

            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1) And OpenStock <> 0 Then

                ws.Cells(RowCounter, 9).Value = ws.Cells(I, 1).Value

                CloseStock = ws.Cells(I, 6).Value

                ws.Cells(RowCounter, 10).Value = CloseStock - OpenStock
                
                ws.Cells(RowCounter, 12).Value = TotalVolume

                TotalVolume = 0
                
                ws.Cells(RowCounter, 11).Value = ws.Cells(RowCounter, 10).Value / OpenStock
                
                ws.Cells(RowCounter, 11).NumberFormat = "0.00%"
                
                OpenStock = ws.Cells(I + 1, 3).Value
                
                RowCounter = RowCounter + 1
                
            ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1) And OpenStock = 0 Then

                ws.Cells(RowCounter, 9).Value = ws.Cells(I, 1).Value

                CloseStock = ws.Cells(I, 6).Value

                ws.Cells(RowCounter, 10).Value = CloseStock
                
                ws.Cells(RowCounter, 12).Value = TotalVolume

                TotalVolume = 0
                
                ws.Cells(RowCounter, 11).Value = ws.Cells(RowCounter, 10).Value
                
                ws.Cells(RowCounter, 11).NumberFormat = "0.00%"
                
                OpenStock = ws.Cells(I + 1, 3).Value
                
                RowCounter = RowCounter + 1
                
            End If
            
        Next I

    For I = 2 To lastrow

        If ws.Cells(I, 10).Value > 0 Then

            ws.Cells(I, 10).Interior.ColorIndex = 4

        ElseIf ws.Cells(I, 10).Value < 0 Then

            ws.Cells(I, 10).Interior.ColorIndex = 3

        End If

    Next I

    Dim Max1, Max2 As Double
    Dim Max3 As LongLong
    Dim Rw1, Rw2, Rw3 As Integer

    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    Max1 = WorksheetFunction.Max(ws.Range("K:K"))
    Max2 = WorksheetFunction.Min(ws.Range("K:K"))
    Max3 = WorksheetFunction.Max(ws.Range("L:L"))

    Rw1 = WorksheetFunction.Match(Max1, ws.Range("K:K"), 0)
    Rw2 = WorksheetFunction.Match(Max2, ws.Range("K:K"), 0)
    Rw3 = WorksheetFunction.Match(Max3, ws.Range("L:L"), 0)

    ws.Cells(2, 15).Value = ws.Cells(Rw1, 9).Value
    ws.Cells(3, 15).Value = ws.Cells(Rw2, 9).Value
    ws.Cells(4, 15).Value = ws.Cells(Rw3, 9).Value

    ws.Cells(2, 16).Value = Max1
    ws.Cells(3, 16).Value = Max2
    ws.Cells(4, 16).Value = Max3
    
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"

    ws.Columns("A:P").AutoFit
    
  Next ws

End Sub

