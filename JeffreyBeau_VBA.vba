Sub StockSummary()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim ticker_name As String

    Dim opening_price As Double

    Dim closing_price As Double

    Dim yearly_change As Double

    Dim percent_change As Double

    Dim current_row As String

    Dim next_row As String

    Dim previous_row As String

    Dim total_stock_volume As Double
    total_stock_volume = 0

    Dim summary_table_row As Integer
    summary_table_row = 2

    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim a As Long
    
    For a = 2 To last_row

        current_row = ws.Cells(a, 1).Value
        next_row = ws.Cells(a + 1, 1).Value
        previous_row = ws.Cells(a - 1, 1).Value
    
        If current_row <> previous_row Then
            opening_price = ws.Cells(a, 3).Value
    
        ElseIf current_row <> next_row Then
    
            ticker_name = current_row
        
            closing_price = ws.Cells(a, 6).Value
        
            yearly_change = closing_price - opening_price
        
            If yearly_change = 0 Or opening_price = 0 Then
                percent_change = 0
            
            Else
                percent_change = yearly_change / opening_price
            
            End If
        
        Debug.Print ticker_name
        
        total_stock_volume = total_stock_volume + ws.Cells(a, 7).Value
        
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("I" & summary_table_row).Value = ticker_name
        
        ws.Range("J1").Value = "Yearly Change"
        
        ws.Range("J" & summary_table_row).Value = yearly_change
        
        ws.Range("K1").Value = "% Change"
        
        ws.Range("K" & summary_table_row).Value = percent_change
        
        ws.Columns("K").NumberFormat = "0.00%"
        
        ws.Range("K" & summary_table_row).Style = "Percent"
        
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("L" & summary_table_row).Value = total_stock_volume
        
        summary_table_row = summary_table_row + 1
        
        yearly_change = 0
        
        percent_change = 0
        
        total_stock_volume = 0
        
        Else
        
            total_stock_volume = total_stock_volume + ws.Cells(a, 7).Value
        
        End If

    Next a

    last_row_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim b As Long

    For b = 2 To last_row_summary

        Dim summary_change As Double
        summary_change = ws.Cells(b, 10).Value

        If summary_change >= 0 Then
            ws.Cells(b, 10).Interior.ColorIndex = 10
        
        Else
            ws.Cells(b, 10).Interior.ColorIndex = 3
    
        End If
    
    Next b

    Dim max As Double
    
    Dim min As Double

    Dim greatest_total_volume As Double
    
    Dim c As Long

    For c = 2 To last_row_summary
    
        Calculate
    
        percent_column = ws.Range("K2:K" & last_row_summary).Value
    
        volume_column = ws.Range("L2:L" & last_row_summary).Value
    
        max = ws.Application.WorksheetFunction.max(percent_column)
    
        min = ws.Application.WorksheetFunction.min(percent_column)
    
        greatest_total_volume = ws.Application.WorksheetFunction.max(volume_column)
    
        ws.Range("O1").Value = "Ticker"
    
        ws.Range("P1").Value = "Value"
    
        ws.Range("N2").Value = "Greatest % Increase"
    
        ws.Range("P2").Value = max
    
        ws.Range("N3").Value = "Greatest % Decrease"
    
        ws.Range("P3").Value = min
    
        ws.Range("N4").Value = "Greatest Total Volume"
    
        ws.Range("P4").Value = greatest_total_volume
    
        ws.Range("P2:P3").Style = "Percent"
    
        ws.Range("P2:P3").NumberFormat = "0.00%"
        
    Next c
    
    Dim d As Long
    
    Dim result As String
    
    For d = 2 To last_row_summary
    
        result = ws.Cells(d, 9).Value
    
        If ws.Cells(d, 11).Value = max Then
            ws.Range("O2").Value = result
        
        ElseIf ws.Cells(d, 11).Value = min Then
            ws.Range("O3").Value = result
        
        ElseIf ws.Cells(d, 12).Value = greatest_total_volume Then
            ws.Range("O4").Value = result
        
    End If
    
    Next d
    
    ws.Columns("A:P").AutoFit

Next ws

End Sub
