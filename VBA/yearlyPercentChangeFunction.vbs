Attribute VB_Name = "yearlyPercentChangeFunction"
'This was similar to year change but tweaked for percent formula
Sub yearlyPercentChange()
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim yearlyPercentChange As Double
    Dim i As Long
    Dim j As Long
    
    'Apply to all worksheets Dim
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ws.Range("K1").Value = "Percent Change"
        
        j = 2
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ticker Then
                If i > 2 Then
                    yearlyChange = closingPrice - openingPrice
                    
                   'Here is the difference, instead of just yearly change you need to now define percentage change formula as well
                    yearlyPercentChange = yearlyChange / openingPrice
                    ws.Cells(j, "K").Value = yearlyPercentChange
                    j = j + 1
                End If
                ticker = ws.Cells(i, "A").Value
                openingPrice = ws.Cells(i, "C").Value
            End If
            
            closingPrice = ws.Cells(i, "F").Value
            
            If i = lastRow Then
                yearlyChange = closingPrice - openingPrice
                yearlyPercentChange = yearlyChange / openingPrice
                ws.Cells(j, "K").Value = yearlyPercentChange
            End If
            
        Next i
    Next ws
End Sub


