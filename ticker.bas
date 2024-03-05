Attribute VB_Name = "Module1"
Sub ticker()

  ' Set an initial variable for holding the brand name
  Dim ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim vol As Double
  vol = 0
  Dim ov As Double
  ov = 0
  Dim cv As Double
  cv = 0
  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
 For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    For i = 2 To lastrow
            If ov = 0 Then
                    ov = Cells(i, 3).Value
                End If
    ' Check if we are still within the same credit card brand, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
                ticker = Cells(i, 1).Value

      ' Add to the Brand Total
                vol = vol + Cells(i, 7).Value
                cv = Cells(i, 6).Value
      ' Print the Credit Card Brand in the Summary Table
                Range("I" & Summary_Table_Row).Value = ticker

      ' Print the Brand Amount to the Summary Table
                Range("J" & Summary_Table_Row).Value = vol
                Range("K" & Summary_Table_Row).Value = cv - ov
                Range("L" & Summary_Table_Row).Value = (cv - ov / ov)

      ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
                vol = 0
                cv = 0
                ov = 0
    ' If the cell immediately following a row is the same brand...
                Else

      ' Add to the Brand Total
                vol = vol + Cells(i, 7).Value
            End If

    Next i
    Next ws

End Sub
