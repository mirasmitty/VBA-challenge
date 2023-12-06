Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
 
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearChange As Double
        Dim Name As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
 
        OpenPrice = Cells(2, Column + 2).Value

        For i = 2 To LastRow

            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        
                Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Name
    
                ClosePrice = Cells(i, Column + 5).Value
          
                YearChange = ClosePrice - OpenPrice
                Cells(Row, Column + 9).Value = YearChange
       
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearChange / OpenPrice
                    Cells(Row, Column + 10).Value = PercentChange
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                
                End If
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                Row = Row + 1
                OpenPrice = Cells(i + 1, Column + 2)
                Volume = 0
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i

        YearChangeLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        For j = 2 To YearChangeLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"

        For Z = 2 To YearChangeLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & YearChangeLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & YearChangeLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & YearChangeLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next ws
        
End Sub


