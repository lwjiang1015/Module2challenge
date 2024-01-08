Attribute VB_Name = "Module1"
Sub Module2()

Dim ws As Worksheet
Dim total As Double
Dim j As Integer
Dim yearlychange As Double
Dim percentchange As Double

For Each ws In Worksheets
    
    total = 0
    j = 0
    yearlychange = 0
    percentchange = 0
    
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ws.Range("I" & j + 2) = ws.Cells(i, 1).Value
        
            ws.Range("J" & j + 2) = total + ws.Cells(i, 7).Value
            
            ws.Range("K" & j + 2) = yearlychage + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
            
            ws.Range("L" & j + 2) = percentchange + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 6).Value
            
        
            total = 0
            j = j + 1
            yearlyhcange = 0
            percentchange = 0
            
        
        Else
            total = total + ws.Cells(i, 7).Value
            yearlychange = yearlychange + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
            percentchange = percentchange + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 6).Value
            
            If yearlychange <= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
            
            End If
            
        End If
        
        
    Next i
            
Next ws
        

End Sub
