Attribute VB_Name = "YearlyChangeFunction"
'I was able to try a different approach, instead of using previous method due to taking awhile to run, it requires the data to be aggregated by tciker symbol and date in descending order...looks like the dataset came this way, but this is to explain what must be true for this code to work properly
Sub yearlyChange()
    
    'This uses similar dim setup, but now you have the full variables of the yearly change formula all stored as variables and counters
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim i As Long
    Dim j As Long
    
    
    'Establishes last row ow dataset...this will loop through as a whole now
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Puts a header name to the Yearly Change column
    ActiveSheet.Range("J1").Value = "Yearly Change"
    
    'This below code now keeps an open counter for both the i row of which you want test and also j column that will eventually print the last value
    j = 2
    For i = 2 To lastRow
    
        'First check to see if a new ticker symbol is found in column A
        If ActiveSheet.Cells(i, "A").Value <> ticker Then
            If i > 2 Then
                
                'This is where the counter for the final yearly change column is established
                yearlyChange = closingPrice - openingPrice
                ActiveSheet.Cells(j, "J").Value = yearlyChange
                
                'If not found, then move to next one
                j = j + 1
            End If
            
            'This next bit will grab the opening price by looping through, to then populate the yearly change calc
            ticker = ActiveSheet.Cells(i, "A").Value
            openingPrice = ActiveSheet.Cells(i, "C").Value
        End If
        
          'This next bit will grab the closing price by looping through, to then populate the yearly change calc
        closingPrice = ActiveSheet.Cells(i, "F").Value
        
       ' Finally the yearly chage gets stored
        If i = lastRow Then
            yearlyChange = closingPrice - openingPrice 'calculate the yearly change
'            ActiveSheet.Cells(j, "K").Value = yearlyChange 'this was use to make sure and test you get the correct ticker values displayed and the aggreation is done properly, it can remain commented out
        End If
        
        ActiveSheet.Cells(j, "J").Value = ticker
    Next i
End Sub
'This calls the same function above but applies it to all sheets via a for loop
Sub ApplyYearlyChangeToAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        'Here is where you call the function
        yearlyChange
    Next ws
End Sub

