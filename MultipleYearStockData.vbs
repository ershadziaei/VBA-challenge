Attribute VB_Name = "Module1"
'Created by Ershad Ziaei (12-20-2022) as VBA Challenge

Sub MultipleYearStockData():
    For Each ws In Worksheets
    
        'Definition
        Dim WorksheetName As String
        Dim Percent As Double
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxVol As Double
        Dim Ticknum As Long
        Dim LastRowAll As Long
        Dim LastRowSort As Long

        WorksheetName = ws.Name

        'Column Title
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'First Data rRow
        Ticknum = 2
        
        'First Try
        j = 2
        
        'Last Data Row
        LastRowAll = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Stock Individual Data
            For i = 2 To LastRowAll
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(Ticknum, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Ticknum, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    If ws.Cells(Ticknum, 10).Value < 0 Then
                    ws.Cells(Ticknum, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(Ticknum, 10).Interior.ColorIndex = 4
                    End If
                    If ws.Cells(j, 3).Value <> 0 Then
                    Percent = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(Ticknum, 11).Value = Format(Percent, "Percent")
                    Else
                    ws.Cells(Ticknum, 11).Value = Format(0, "Percent")
                    End If
                ws.Cells(Ticknum, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                Ticknum = Ticknum + 1
                j = i + 1
                End If
            Next i
            
        'Stock Summary Data
        LastRowSort = ws.Cells(Rows.Count, 9).End(xlUp).Row
        MaxVol = ws.Cells(2, 12).Value
        MaxIncrease = ws.Cells(2, 11).Value
        MaxDecrease = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowSort
                If ws.Cells(i, 12).Value > MaxVol Then
                MaxVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxVol = MaxVol
                End If
                If ws.Cells(i, 11).Value > MaxIncrease Then
                MaxIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxIncrease = MaxIncrease
                End If
                If ws.Cells(i, 11).Value < MaxDecrease Then
                MaxDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxDecrease = MaxDecrease
                End If
            ws.Cells(2, 17).Value = Format(MaxIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(MaxDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(MaxVol, "Scientific")
            Next i
    Next ws
End Sub
