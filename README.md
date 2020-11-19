# VBA-Homework

Sub stocks()
Dim ws As Worksheet
    
    For Each ws In Worksheets

Dim ticker As String
Dim total_volume As Double

Dim year_open As Double
Dim year_close As Double

Dim summary_row As Integer
summary_row = 2
    

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Toal Volume"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"

Dim last_row As Double
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row
        
        If (ws.Cells(i, 3).Value = 0) Then
              If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then

            ticker = ws.Cells(i, 1).Value
         End If
              
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                If (ws.Cells(i - 1, 1) <> ws.Cells(i, 1).Value) Then
                year_open = ws.Cells(i, 3).Value
                
          End If
                
        Else
            
        ticker = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        year_close = ws.Cells(i, 6).Value
        
        ws.Cells(summary_row, 9).Value = ticker
        ws.Cells(summary_row, 10).Value = total_volume
        
        If (total_volume > 0) Then
         ws.Cells(summary_row, 11).Value = year_close - year_open
        
        If (ws.Cells(summary_row, 11).Value > 0) Then
          ws.Cells(summary_row, 11).Interior.ColorIndex = 4
          
        Else
          ws.Cells(summary_row, 11).Interior.ColorIndex = 3
          
          End If
          
          ws.Cells(summary_row, 12).Value = ws.Cells(summary_row, 11).Value / year_open
          
          Else
           ws.Cells(summary_row, 11).Value = 0
           ws.Cells(summary_row, 12).Value = 0
           End If
           
           ws.Cells(summary_row, 12).Style = "percent"
           total_volume = 0
           summary_row = summary_row + 1
           
           End If
           Next i
           Next ws
End Sub


