Attribute VB_Name = "Module1"

Sub tickerStock()
    
    Dim ws As Worksheet
    Dim last_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim quaterly_change As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim row As Long
    Dim column As Integer
    Dim ticker As String
    Dim j As Long
    Dim k As Long
    Dim quaterly_change_last_row As Long
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    Dim max_increase_value As Double
    Dim max_decrease_value As Double
    Dim max_volume_value As Double
    
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        
            ws.Activate
            
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
            
            If ws.Cells(1, 9).Value <> "Ticker" Then
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Quarterly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
            End If
            
            open_price = ws.Cells(2, 3).Value
            row = 2
            volume = 0
            
            For i = 2 To last_row
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Cells(row, 9).Value = ticker
                    
                    close_price = ws.Cells(i, 6).Value
                    
                    quaterly_change = close_price - open_price
                    ws.Cells(row, 10).Value = quaterly_change
                    
                    If quaterly_change <> 0 Then
                        percent_change = quaterly_change / open_price
                    Else
                        percent_change = 0
                    End If
                    ws.Cells(row, 11).Value = percent_change
                    ws.Cells(row, 11).NumberFormat = "0.00%"
                    
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Cells(row, 12).Value = volume
                    
                    row = row + 1
                    open_price = ws.Cells(i + 1, 3).Value
                    volume = 0
                Else
                    volume = volume + ws.Cells(i, 7).Value
                End If
            Next i
            
            quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
            For j = 2 To quaterly_change_last_row
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Bright green
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red for negative changes
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = xlNone ' No color for zero changes
                End If
            Next j
            
            max_increase_value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row))
            max_decrease_value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row))
            max_volume_value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row))
            
            For k = 2 To quaterly_change_last_row
                If ws.Cells(k, 11).Value = max_increase_value Then
                    max_increase_ticker = ws.Cells(k, 9).Value
                    ws.Cells(2, 16).Value = max_increase_ticker
                    ws.Cells(2, 17).Value = max_increase_value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(k, 11).Value = max_decrease_value Then
                    max_decrease_ticker = ws.Cells(k, 9).Value
                    ws.Cells(3, 16).Value = max_decrease_ticker
                    ws.Cells(3, 17).Value = max_decrease_value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(k, 12).Value = max_volume_value Then
                    max_volume_ticker = ws.Cells(k, 9).Value
                    ws.Cells(4, 16).Value = max_volume_ticker
                    ws.Cells(4, 17).Value = max_volume_value
                End If
            Next k
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
         
            
            ws.Range("I:Q").EntireColumn.AutoFit
        
    Next ws
    
    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets(1).Select

End Sub






