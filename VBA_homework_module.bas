Attribute VB_Name = "Module2"
Sub VBA_homework_loop():
    Dim sht As Worksheet
    Dim last_row As String
    
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Variant
    Dim table_row As Integer
    Dim date1 As Long
    
    For Each sht In Worksheets
        sht.Range("I1").Value = "Ticker"
        sht.Range("L1").Value = "Total Stock Volume"
        
        total_stock = 0
        table_row = 2
    
        
        last_row = sht.Cells.SpecialCells(xlCellTypeLastCell).Row
        
        For i = 2 To last_row
            If sht.Cells(i + 1, 1).Value <> sht.Cells(i, 1).Value Then
                ticker = sht.Cells(i, 1).Value
                total_stock = total_stock + sht.Cells(i, 7).Value
                sht.Range("I" & table_row).Value = ticker
                sht.Range("L" & table_row).Value = total_stock
                table_row = table_row + 1
                total_stock = 0
            Else
                total_stock = total_stock + sht.Cells(i, 7).Value
            End If
        Next i
    Next sht
    
End Sub
