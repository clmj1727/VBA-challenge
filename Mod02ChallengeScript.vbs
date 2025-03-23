VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickerStock()
    Dim ws As Worksheet
    Application.ScreenUpdating = False ' Speeds up execution

    ' Loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ' Find the last row in column A
        Dim last_row As Long
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables
        Dim open_price As Double
        Dim close_price As Double
        Dim quarterly_change As Double
        Dim ticker As String
        Dim percent_change As Double
        Dim volume As Double
        Dim row As Integer
        Dim column As Integer
        Dim i As Long
        
        volume = 0
        row = 2
        column = 1
        
        ' Set the initial open price
        open_price = ws.Cells(2, column + 2).Value
        
        ' Loop through tickers
        For i = 2 To last_row
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ' Capture ticker name
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, column + 8).Value = ticker
                
                ' Capture close price
                close_price = ws.Cells(i, column + 5).Value
                
                ' Calculate quarterly change
                quarterly_change = close_price - open_price
                ws.Cells(row, column + 9).Value = quarterly_change
                
                ' Calculate percent change
                If open_price <> 0 Then
                    percent_change = quarterly_change / open_price
                Else
                    percent_change = 0
                End If
                
                ws.Cells(row, column + 10).Value = percent_change
                ws.Cells(row, column + 10).NumberFormat = "0.00%"
                
                ' Calculate total volume
                volume = volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, column + 11).Value = volume
                
                ' Move to the next row
                row = row + 1
                
                ' Reset open price for next ticker
                open_price = ws.Cells(i + 1, column + 2).Value
                volume = 0 ' Reset volume
                
            Else
                ' Continue accumulating volume
                volume = volume + ws.Cells(i, column + 6).Value
            End If
        Next i
        
        ' Find the last row in the ticker column
        Dim quarterly_change_last_row As Long
        quarterly_change_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        
        ' Apply color formatting for quarterly change
        Dim j As Long
        For j = 2 To quarterly_change_last_row
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10 ' Green
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red
            End If
        Next j
        
        ' Set headers for greatest increase, decrease, and volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Identify highest values
        Dim k As Long
        Dim max_increase As Double, min_decrease As Double, max_volume As Double
        Dim max_ticker As String, min_ticker As String, max_vol_ticker As String
        
        max_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row))
        min_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row))
        max_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row))
        
        ' Loop to find corresponding tickers
        For k = 2 To quarterly_change_last_row
            If ws.Cells(k, 11).Value = max_increase Then
                max_ticker = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 11).Value = min_decrease Then
                min_ticker = ws.Cells(k, 9).Value
            ElseIf ws.Cells(k, 12).Value = max_volume Then
                max_vol_ticker = ws.Cells(k, 9).Value
            End If
        Next k
        
        ' Assign results
        ws.Cells(2, 16).Value = max_ticker
        ws.Cells(2, 17).Value = max_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = min_ticker
        ws.Cells(3, 17).Value = min_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = max_vol_ticker
        ws.Cells(4, 17).Value = max_volume
        
        ' Format columns
        ws.Range("I:Q").Font.Bold = True
        ws.Range("I:Q").EntireColumn.AutoFit
    Next ws

    ' Select "Q1" if it exists
    On Error Resume Next
    Worksheets("Q1").Select
    On Error GoTo 0
    
    Application.ScreenUpdating = True ' Restore screen updating
End Sub

