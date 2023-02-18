Attribute VB_Name = "UniqueValues"
'This is to name the function
Sub Ticker_uniqueValues()
    'Sets a dim for worksheets so this code applies to all sheets
    Dim ws As Worksheet
    
    'Sets other dims for base function
    Dim lastRow As Long
    Dim cell As Range
    Dim UniqueValues As New Collection
    Dim i As Long
    
    
    'Sets loop to loop through each sheet, apply a code, then goes to next sheet
    For Each ws In ActiveWorkbook.Worksheets
    
        'Next, this places the given worksheet's cell value Ticker in J1 to name your new column
        ws.Range("I1").Value = "Ticker"
        
        ' This will establish/find the last row in the ticker column as a variable to use in for loop
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        ' Now this for loop will go through each cell in column A of the ticker column to only only pull the unique values
        For Each cell In ws.Range("A2:A" & lastRow)
            On Error Resume Next
            UniqueValues.Add cell.Value, CStr(cell.Value)
            On Error GoTo 0
        Next cell
        
        
        'Finally, this will print the unique values to the new Ticker column J
        For i = 1 To UniqueValues.Count
            ws.Cells(1 + i, "I").Value = UniqueValues(i)
        Next i
        
        'In order to reset the counters for the next worksheet you must clear the collection values you've already stored
        Set UniqueValues = New Collection
        
    Next ws

End Sub





