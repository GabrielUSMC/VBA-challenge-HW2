Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim TickerName As String
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim TVol As Variant
    Dim OutputIndex As Long
    Dim MaxPerInc As Double
    Dim MaxPerInc_name As String
    Dim MaxPerDec As Double
    Dim MaxPerDec_name As String
    Dim MaxVol As Variant
    Dim MaxVol_name As String

    For Each ws In Worksheets

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        MaxPerInc = -1000.01
        MaxPerDec = 1000.01
        MaxVol = -1000
        
        ws.Cells(1, 10) = ws.Cells(1, 1)
        ws.Cells(1, 11) = "<Yearly Change>"
        ws.Cells(1, 12) = "<Percent Change>"
        ws.Cells(1, 13) = "<Total Volume>"
        OutputIndex = 2
        StockOpen = ws.Cells(2, 3)
        ws.Cells(OutputIndex, 10) = ws.Cells(2, 1)
        TVol = ws.Cells(2, 7)
        
        For Row = 3 To LastRow + 1
            If (ws.Cells(Row, 1) = ws.Cells(Row - 1, 1)) Then
                TVol = TVol + ws.Cells(Row, 7).Value
            
            Else
                StockClose = ws.Cells(Row - 1, 6)
                ws.Cells(OutputIndex, 11) = StockClose - StockOpen
                If StockOpen <> 0 Then
                   ws.Cells(OutputIndex, 12) = ws.Cells(OutputIndex, 11) / StockOpen
                End If
                If ws.Cells(OutputIndex, 11) > 0 Then
                    ws.Cells(OutputIndex, 11).Interior.ColorIndex = 4
                    ws.Cells(OutputIndex, 12).Interior.ColorIndex = 4
                Else
                    ws.Cells(OutputIndex, 11).Interior.ColorIndex = 3
                    ws.Cells(OutputIndex, 12).Interior.ColorIndex = 3
                End If
                
                ws.Cells(OutputIndex, 12).NumberFormat = "0.00%"
                ws.Cells(OutputIndex, 13) = TVol
                TVol = ws.Cells(Row, 7)
                
                If ws.Cells(OutputIndex, 12) > MaxPerInc Then
                    MaxPerInc = ws.Cells(OutputIndex, 12)
                    MaxPerInc_name = ws.Cells(OutputIndex, 10)
                End If
                If ws.Cells(OutputIndex, 12) < MaxPerDec Then
                    MaxPerDec = ws.Cells(OutputIndex, 12)
                    MaxPerDec_name = ws.Cells(OutputIndex, 10)
                End If
                If ws.Cells(OutputIndex, 13) > MaxVol Then
                    MaxVol = ws.Cells(OutputIndex, 13)
                    MaxVol_name = ws.Cells(OutputIndex, 10)
                End If
                
                StockOpen = ws.Cells(Row, 3)
                OutputIndex = OutputIndex + 1
                ws.Cells(OutputIndex, 10) = ws.Cells(Row, 1)
            End If
        
        Next Row
        
        ws.Range("P1") = "<ticker>"
        ws.Range("Q1") = "<value>"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("P2") = MaxPerInc_name
        ws.Range("Q2") = MaxPerInc
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("O3") = "Greatest % Decrease"
         ws.Range("P3") = MaxPerDec_name
        ws.Range("Q3") = MaxPerDec
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("O4") = "Greastest total volume"
         ws.Range("P4") = MaxVol_name
        ws.Range("Q4") = MaxVol
        
    Next ws

End Sub
