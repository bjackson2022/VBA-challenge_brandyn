Attribute VB_Name = "SumIFS"
'This is to name the function for the Total Stock Volume
Sub SumIfsLoop()

    'Apply to all worksheets Dim
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    Dim criteriaRange As Range
    Dim sumRange As Range
    Dim resultRange As Range
    
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        'This is to name the header column for the Total Stock Volume calc
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' This sets the counts of the last row of each criteria in sumifs formula
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        lastRow2 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastRow3 = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
        
        
        ' This sets the range values for the criteria and sum ranges
        Set criteriaRange = ws.Range("A2:A" & lastRow2)
        Set sumRange = ws.Range("G2:G" & lastRow3)
        
        ' This sets the result range for your new column to store thevalues
        Set resultRange = ws.Range("L2:L" & lastRow)

        ' This is the for each time you need to run the SUMIFS function
        For i = 1 To resultRange.Cells.Count
            resultRange.Cells(i).Value = WorksheetFunction.SumIFS(sumRange, criteriaRange, criteriaRange.Cells(i, "I"))
        Next i
    
    Next ws

End Sub
