Attribute VB_Name = "Module1"
Sub ticker()

  Dim ticker As String
  Dim vol As Double
  vol = 0
  Dim ov As Double
  ov = 0
  Dim cv As Double
  cv = 0
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
    For Each ws In Worksheets 'refuses too run other sheets
            
            lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
            For i = 2 To lastrow
                If ov = 0 Then
                        ov = Cells(i, 3).Value
                End If
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    ticker = Cells(i, 1).Value
                    vol = vol + Cells(i, 7).Value
                    cv = Cells(i, 6).Value
                    
                    Range("I1").Value = "Ticker"
                    Range("I" & Summary_Table_Row).Value = ticker
                    Range("J1").Value = "Total Stock volume"
                    Range("J" & Summary_Table_Row).Value = vol
                    Range("K1").Value = "Yearly Change"
                    Range("K" & Summary_Table_Row).Value = cv - ov
                    Range("L1").Value = "Percent Change"
                    Range("L" & Summary_Table_Row).Value = ((cv - ov) / ov)
                    Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    If Range("K" & Summary_Table_Row).Value < 0# Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    If Range("L" & Summary_Table_Row).Value < 0# Then
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If

                    Summary_Table_Row = Summary_Table_Row + 1
      
                    vol = 0
                    cv = 0
                    ov = 0

                    Else

                    vol = vol + Cells(i, 7).Value
                End If
            
                Increase = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_Row))
                Decrease = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_Row))
                Tvol = Application.WorksheetFunction.Max(Range("J2:J" & Summary_Table_Row))
            
                Cells(1, 15).Value = "Ticker"
                Cells(1, 16).Value = "Value"
            
                If Cells(i, 11).Value = Increase Then
                    Cells(2, 16).Value = Cells(i, 12).Value
                    Cells(2, 16).NumberFormat = "0.00%"
                    Cells(2, 15).Value = Cells(i, 9).Value
                    Cells(2, 14).Value = "Greatest Increase"
                End If
                If Cells(i, 11).Value = Decrease Then
                    Cells(3, 16).Value = Cells(i, 12).Value
                    Cells(3, 16).NumberFormat = "0.00%"
                    Cells(3, 15).Value = Cells(i, 9).Value
                    Cells(3, 14).Value = "Greatest Decrease"
                
                End If
                If Cells(i, 10).Value = Tvol Then
                    Cells(4, 16).Value = Cells(i, 10).Value
                    Cells(4, 16).NumberFormat = "general"
                    Cells(4, 15).Value = Cells(i, 9).Value
                    Cells(4, 14).Value = "Greatest Total Volume"
                End If
            Next i
    'MsgBox (ws.Name)
    Next ws
Run "ticker19"
Run "ticker20"
End Sub
Sub ticker19()
    Worksheets("2019").Activate
    Dim ticker As String
    Dim vol As Double
    vol = 0
    Dim ov As Double
    ov = 0
    Dim cv As Double
    cv = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    For Each ws In Worksheets 'refuses too run other sheets
            
            lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
            For i = 2 To 4000 '4000 to cut down on processing time
                If ov = 0 Then
                        ov = Cells(i, 3).Value
                End If
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    ticker = Cells(i, 1).Value
                    vol = vol + Cells(i, 7).Value
                    cv = Cells(i, 6).Value
                    
                    Range("I1").Value = "Ticker"
                    Range("I" & Summary_Table_Row).Value = ticker
                    Range("J1").Value = "Total Stock volume"
                    Range("J" & Summary_Table_Row).Value = vol
                    Range("K1").Value = "Yearly Change"
                    Range("K" & Summary_Table_Row).Value = cv - ov
                    Range("L1").Value = "Percent Change"
                    Range("L" & Summary_Table_Row).Value = ((cv - ov) / ov)
                    Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    If Range("K" & Summary_Table_Row).Value < 0# Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    If Range("L" & Summary_Table_Row).Value < 0# Then
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If

                    Summary_Table_Row = Summary_Table_Row + 1
      
                    vol = 0
                    cv = 0
                    ov = 0

                    Else

                    vol = vol + Cells(i, 7).Value
                End If
            
                Increase = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_Row))
                Decrease = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_Row))
                Tvol = Application.WorksheetFunction.Max(Range("J2:J" & Summary_Table_Row))
            
                Cells(1, 15).Value = "Ticker"
                Cells(1, 16).Value = "Value"
            
                If Cells(i, 11).Value = Increase Then
                    Cells(2, 16).Value = Cells(i, 12).Value
                    Cells(2, 15).Value = Cells(i, 9).Value
                    Cells(2, 14).Value = "Greatest Increase"
                End If
                If Cells(i, 11).Value = Decrease Then
                    Cells(3, 16).Value = Cells(i, 12).Value
                    Cells(3, 15).Value = Cells(i, 9).Value
                    Cells(3, 14).Value = "Greatest Decrease"
                End If
                If Cells(i, 10).Value = Tvol Then
                    Cells(4, 16).Value = Cells(i, 10).Value
                    Cells(4, 15).Value = Cells(i, 9).Value
                    Cells(4, 14).Value = "Greatest Total Volume"
                End If
            Next i
    'MsgBox (ws.Name)
    Next ws
End Sub
Sub ticker20()
  Worksheets("2020").Activate
  Dim ticker As String
  Dim vol As Double
  vol = 0
  Dim ov As Double
  ov = 0
  Dim cv As Double
  cv = 0
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
    For Each ws In Worksheets 'refuses too run other sheets
            
            lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
            For i = 2 To 4000 '4000 to cut down on processing time
                If ov = 0 Then
                        ov = Cells(i, 3).Value
                End If
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    ticker = Cells(i, 1).Value
                    vol = vol + Cells(i, 7).Value
                    cv = Cells(i, 6).Value
                    
                    Range("I1").Value = "Ticker"
                    Range("I" & Summary_Table_Row).Value = ticker
                    Range("J1").Value = "Total Stock volume"
                    Range("J" & Summary_Table_Row).Value = vol
                    Range("K1").Value = "Yearly Change"
                    Range("K" & Summary_Table_Row).Value = cv - ov
                    Range("L1").Value = "Percent Change"
                    Range("L" & Summary_Table_Row).Value = ((cv - ov) / ov)
                    Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    If Range("K" & Summary_Table_Row).Value < 0# Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    If Range("L" & Summary_Table_Row).Value < 0# Then
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If

                    Summary_Table_Row = Summary_Table_Row + 1
      
                    vol = 0
                    cv = 0
                    ov = 0

                    Else

                    vol = vol + Cells(i, 7).Value
                End If
            
                Increase = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_Row))
                Decrease = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_Row))
                Tvol = Application.WorksheetFunction.Max(Range("J2:J" & Summary_Table_Row))
            
                Cells(1, 15).Value = "Ticker"
                Cells(1, 16).Value = "Value"
            
                If Cells(i, 11).Value = Increase Then
                    Cells(2, 16).Value = Cells(i, 12).Value
                    Cells(2, 15).Value = Cells(i, 9).Value
                    Cells(2, 14).Value = "Greatest Increase"
                End If
                If Cells(i, 11).Value = Decrease Then
                    Cells(3, 16).Value = Cells(i, 12).Value
                    Cells(3, 15).Value = Cells(i, 9).Value
                    Cells(3, 14).Value = "Greatest Decrease"
                End If
                If Cells(i, 10).Value = Tvol Then
                    Cells(4, 16).Value = Cells(i, 10).Value
                    Cells(4, 15).Value = Cells(i, 9).Value
                    Cells(4, 14).Value = "Greatest Total Volume"
                End If
            Next i
    'MsgBox (ws.Name)
    Next ws
End Sub
