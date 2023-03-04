Attribute VB_Name = "overallAnalysis"
Sub stockAnalysis()

    'Declare variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
        'Declare variables for greatest increase, decrease, and volume
        Dim maxPercentIncreaseTicker As String
        Dim maxPercentDecreaseTicker As String
        Dim maxVolumeTicker As String
        Dim maxPercentIncrease As Double
        Dim maxPercentDecrease As Double
        Dim maxVolume As Double
        
        'Initialize variables for greatest increase, decrease, and volume
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxVolume = 0
    
        'Initialize output row variable
        Dim j As Integer
        j = 2
        
        'Find the last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through each row of data
        For i = 2 To lastRow
        
            'Check if we have moved on to a new stock ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Get the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                'Get the opening price for the year
                openingPrice = ws.Cells(i - 11, 3).Value
                
                'Get the closing price for the year
                closingPrice = ws.Cells(i, 6).Value
                
                'Calculate the yearly change
                yearlyChange = closingPrice - openingPrice
                
                'Calculate the percent change
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                'Get the total stock volume
                totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(i - 11, 7), ws.Cells(i, 7)))
                
                'Output the results
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Volume"
                
                ws.Range("I" & j).Value = ticker
                ws.Range("J" & j).Value = yearlyChange
                ws.Range("K" & j).Value = percentChange
                ws.Range("K" & j).NumberFormat = "0.00%"
                ws.Range("L" & j).Value = totalVolume
                
                'Update variables for greatest increase, decrease, and volume
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
                'Move to the next row for output
                j = j + 1
                
            End If
        
        Next i
        
        'Output the results for greatest increase, decrease, and volume
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = maxPercentIncreaseTicker
        ws.Range("P3").Value = maxPercentDecreaseTicker
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q2").Value = maxPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = maxPercentDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = maxVolume
    Next ws
    
End Sub



