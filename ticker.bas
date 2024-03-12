Attribute VB_Name = "Module1"
Sub ticker()
    For Each ws In Worksheets
            Dim ticker As String
            Dim vol As Double
            vol = 0
            Dim ov As Double
            ov = 0
            Dim cv As Double
            cv = 0
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
            For i = 2 To lastrow
                If ov = 0 Then
                        ov = ws.Cells(i, 3).Value
                End If
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    vol = vol + ws.Cells(i, 7).Value
                    cv = ws.Cells(i, 6).Value
                    
                    ws.Range("I1").Value = "Ticker"
                    ws.Range("I" & Summary_Table_Row).Value = ticker
                    ws.Range("J1").Value = "Total Stock volume"
                    ws.Range("J" & Summary_Table_Row).Value = vol
                    ws.Range("K1").Value = "Yearly Change"
                    ws.Range("K" & Summary_Table_Row).Value = cv - ov
                    ws.Range("L1").Value = "Percent Change"
                    ws.Range("L" & Summary_Table_Row).Value = ((cv - ov) / ov)
                    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    If ws.Range("K" & Summary_Table_Row).Value < 0# Then
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    If ws.Range("L" & Summary_Table_Row).Value < 0# Then
                        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If

                    Summary_Table_Row = Summary_Table_Row + 1
      
                    vol = 0
                    cv = 0
                    ov = 0
                    Else

                    vol = vol + ws.Cells(i, 7).Value
                End If
            Next i
            For i = 2 To 3001
                Increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row))
                Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row))
                Tvol = Application.WorksheetFunction.Max(ws.Range("J2:J" & Summary_Table_Row))
            
                ws.Cells(1, 15).Value = "Ticker"
                ws.Cells(1, 16).Value = "Value"
            
                If ws.Cells(i, 11).Value = Increase Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 12).Value
                    ws.Cells(2, 16).NumberFormat = "0.00%"
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 14).Value = "Greatest Increase"
                End If
                If ws.Cells(i, 11).Value = Decrease Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 12).Value
                    ws.Cells(3, 16).NumberFormat = "0.00%"
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 14).Value = "Greatest Decrease"
                
                End If
                If ws.Cells(i, 10).Value = Tvol Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
                    ws.Cells(4, 16).NumberFormat = "general"
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 14).Value = "Greatest Total Volume"
                End If
            Next i
    Next ws
End Sub
