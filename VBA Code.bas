Attribute VB_Name = "Module1"
Sub vba_count_rows_with_data()
'define Variables

Dim iStocks As Long
Dim i As Range
Dim Tricker As String
Dim new_tricker As Integer
Dim TotalStockVolume As Double
Dim DayOneTricker As Double
Dim Sheets As Integer

'Loop to all worksheets

For Each ws In Worksheets

'Summary Name columns name
ws.Cells(1, 9).Value = "Tricker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Tricker = ws.Cells(2, 1)
DayOneOpening = ws.Cells(2, 3)
iStocks = 2
new_tricker = 2
TotalStockVolume = 0

    'loop through each row from the used range
    For Each i In ws.Rows

        'check if the row contains a cell with a value
        If Application.CountA(i) > 0 Then
       DayOneTricker = 2
            'counts the number of rows non-empty Cells
              If ws.Cells(iStocks, 1).Value = Tricker Then
                  TotalStockVolume = TotalStockVolume + ws.Cells(iStocks, 7).Value
                Else
                    Tricker = ws.Cells(iStocks, 1).Value
                    ws.Cells(new_tricker, 9) = ws.Cells(iStocks - 1, 1).Value
                    ws.Cells(new_tricker, 10) = ws.Cells(iStocks - 1, 6).Value - DayOneOpening
                    ws.Cells(new_tricker, 11) = FormatPercent((ws.Cells(new_tricker, 10).Value / DayOneOpening))
                        If ws.Cells(new_tricker, 10) < 0 Then
                        ws.Cells(new_tricker, 10).Interior.ColorIndex = 3
                        Else
                        ws.Cells(new_tricker, 10).Interior.ColorIndex = 4
                        End If
                            ws.Cells(new_tricker, 12) = TotalStockVolume
                            TotalStockVolume = 0
                            DayOneOpening = ws.Cells(iStocks, 3).Value
                            new_tricker = new_tricker + 1
            End If
           iStocks = iStocks + 1

        End If

    Next

Next ws

End Sub

Sub Greatestvalue()
'Loop all sheets

For Each ws In Worksheets
ws.Cells(1, 15).Value = "Tricker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest% Increase"
ws.Cells(3, 14).Value = "Greatest% Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

'find Greatest % increase,decrease & Total Volume

ws.Cells(2, 16).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("K:K")))
ws.Cells(3, 16).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("K:K")))
ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))

GreatestIncrease = ws.Cells(2, 16).Value
GreatestDecrease = ws.Cells(3, 16).Value
GreatestTotal = ws.Cells(4, 16).Value
'Loop to find tricker name

For i = 2 To 3001

    If ws.Cells(i, 11) = GreatestIncrease Then
    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
    
         ElseIf ws.Cells(i, 11) = GreatestDecrease Then
    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
    
        ElseIf ws.Cells(i, 12) = GreatestTotal Then
    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
           
    End If
   
Next i
Next ws

End Sub

