Attribute VB_Name = "Module1"


Sub test()

Dim open_price As Double
Dim close_price As Double
Dim price_change As Double
Dim price_change_percent As Double

Dim greatest_increase As Double
Dim greatest_increase_ticker As String
greatest_increase = 0
Dim greatest_decrease As Double
Dim greatest_decrease_ticker As String
greatest_decrease = 0
Dim greatest_volume As Double
Dim greatest_volume_ticker As String
greatest_volume = 0

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Open Price"
    ws.Range("K1").Value = "Close Price"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim ticker_index As Integer
    
    Dim volume_total As Double
    volume_total = 0
    
    Dim row_num As Long
    For row_num = 2 To last_row
    
        If row_num = 2 Then
            ticker_index = 2
            current_ticker = ws.Cells(row_num, 1).Value
            ws.Cells(ticker_index, 9).Value = current_ticker
            open_price = ws.Cells(row_num, 3).Value
            ws.Cells(ticker_index, 10).Value = open_price
            volume_total = volume_total + ws.Cells(row_num, 7).Value
        ElseIf row_num = last_row Then
            volume_total = volume_total + ws.Cells(row_num, 7).Value
            close_price = ws.Cells(row_num, 6).Value
            ws.Cells(ticker_index, 11).Value = close_price
            price_change = close_price - open_price
            ws.Cells(ticker_index, 12).Value = price_change
            price_change_percent = ((close_price - open_price) / open_price) * 100
            ws.Cells(ticker_index, 13).Value = price_change_percent
            ws.Cells(ticker_index, 14).Value = volume_total
            If price_change > 0 Then
                ws.Cells(ticker_index, 12).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_index, 12).Interior.ColorIndex = 3
            End If
            If price_change_percent > greatest_increase Then
                greatest_increase = price_change_percent
                greatest_increase_ticker = current_ticker
            End If
            If price_change_percent < greatest_decrease Then
                greatest_decrease = price_change_percent
                greatest_decrease_ticker = current_ticker
            End If
            If volume_total > greatest_volume Then
                greatest_volume = volume_total
                greatest_volume_ticker = current_ticker
            End If
        ElseIf ws.Cells(row_num, 1).Value <> current_ticker Then
            volume_total = volume_total + ws.Cells(row_num - 1, 7).Value
            close_price = ws.Cells(row_num - 1, 6).Value
            ws.Cells(ticker_index, 11).Value = close_price
            price_change = close_price - open_price
            ws.Cells(ticker_index, 12).Value = price_change
            price_change_percent = ((close_price - open_price) / open_price) * 100
            ws.Cells(ticker_index, 13).Value = price_change_percent
            ws.Cells(ticker_index, 14).Value = volume_total
            If price_change > 0 Then
                ws.Cells(ticker_index, 12).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_index, 12).Interior.ColorIndex = 3
            End If
            If price_change_percent > greatest_increase Then
                greatest_increase = price_change_percent
                greatest_increase_ticker = current_ticker
            End If
            If price_change_percent < greatest_decrease Then
                greatest_decrease = price_change_percent
                greatest_decrease_ticker = current_ticker
            End If
            If volume_total > greatest_volume Then
                greatest_volume = volume_total
                greatest_volume_ticker = current_ticker
            End If
            ticker_index = ticker_index + 1
            volume_total = ws.Cells(row_num, 7).Value
            current_ticker = ws.Cells(row_num, 1).Value
            ws.Cells(ticker_index, 9).Value = current_ticker
            open_price = ws.Cells(row_num, 3).Value
            ws.Cells(ticker_index, 10).Value = open_price
        Else
            volume_total = volume_total + ws.Cells(row_num, 7).Value
        
        End If
        
        Next row_num
        
    Next ws
    
Worksheets("2018").Range("Q2").Value = "Greatest % Increase"
Worksheets("2018").Range("Q3").Value = "Greatest % Decrease"
Worksheets("2018").Range("Q4").Value = "Greatest Volume"

Worksheets("2018").Range("R1").Value = "Ticker"
Worksheets("2018").Range("R2").Value = greatest_increase_ticker
Worksheets("2018").Range("R3").Value = greatest_decrease_ticker
Worksheets("2018").Range("R4").Value = greatest_volume_ticker

Worksheets("2018").Range("S1").Value = "Value"
Worksheets("2018").Range("S2").Value = greatest_increase
Worksheets("2018").Range("S3").Value = greatest_decrease
Worksheets("2018").Range("s4").Value = greatest_volume


End Sub

