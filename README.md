# VBA-challenge

Sub StockSummary()

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
last_row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row

    current_row = Cells(i, 1).Value
    next_row = Cells(i + 1, 1).Value
    previous_row = Cells(i - 1, 1).Value
    
    If current_row <> previous_row Then
        opening_price = Cells(i, 3).Value
    
    ElseIf current_row <> next_row Then
    
        ticker_name = current_row
        
        closing_price = Cells(i, 6).Value
        
        yearly_change = closing_price - opening_price
        
        If yearly_change = 0 Or opening_price = 0 Then
            percent_change = 0
            
            Else
            percent_change = yearly_change / opening_price
            
        End If
        
        Debug.Print ticker_name
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        Range("I1").Value = "Ticker"
        
        Range("I" & summary_table_row).Value = ticker_name
        
        Range("J1").Value = "Yearly Change"
        
        Range("J" & summary_table_row).Value = yearly_change
        
        Range("K1").Value = "% Change"
        
        Range("K" & summary_table_row).Value = percent_change
        
        Columns("K").NumberFormat = "0.00%"
        
        Range("K" & summary_table_row).Style = "Percent"
        
        Range("L1").Value = "Total Stock Volume"
        
        Range("L" & summary_table_row).Value = total_stock_volume
        
        summary_table_row = summary_table_row + 1
        
        yearly_change = 0
        
        percent_change = 0
        
        total_stock_volume = 0
        
        Else
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
    End If

Next i

last_row_summary = Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To last_row_summary

    Dim summary_change As Double
    summary_change = Cells(j, 10).Value

    If summary_change >= 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 10
        
        Else
        
        Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

Dim max As Double
    
Dim min As Double

Dim greatest_total_volume As Double

For i = 2 To last_row_summary
    
    Calculate
    
    percent_column = Range("K2:K" & last_row_summary).Value
    
    volume_column = Range("L2:L" & last_row_summary).Value
    
    max = Application.WorksheetFunction.max(percent_column)
    
    min = Application.WorksheetFunction.min(percent_column)
    
    greatest_total_volume = Application.WorksheetFunction.max(volume_column)
    
    Range("O1").Value = "Ticker"
    
    Range("P1").Value = "Value"
    
    Range("N2").Value = "Greatest % Increase"
    
    Range("P2").Value = max
    
    Range("N3").Value = "Greatest % Decrease"
    
    Range("P3").Value = min
    
    Range("N4").Value = "Greatest Total Volume"
    
    Range("P4").Value = greatest_total_volume
    
    Range("P2:P3").Style = "Percent"
    
    Range("P2:P3").NumberFormat = "0.00%"

Next i

End Sub
