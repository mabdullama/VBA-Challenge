Sub StockAnalysis()

    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Long
    
    ' loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
        ' set initial values for summary table row and total volume
        summaryTableRow = 1
        totalVolume = 0
        
        ' set headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' find last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' loop through all rows of data
        For i = 2 To lastRow
        
            ' check if we are still looking at the same ticker
            If ws.Cells(i, 1).Value <> ticker Then
                
                ' set ticker and opening price
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                
            End If
            
            ' add volume to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' check if we have moved on to a new ticker
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                ' set closing price and calculate yearly and percent change
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                
                ' output results to summary table
                ws.Range("I" & summaryTableRow + 1).Value = ticker
                ws.Range("J" & summaryTableRow + 1).Value = yearlyChange
                ws.Range("K" & summaryTableRow + 1).Value = percentChange
                ws.Range("L" & summaryTableRow + 1).Value = totalVolume
                
                ' format percent change as a percentage
                ws.Range("K" & summaryTableRow + 1).NumberFormat = "0.00%"
                
                ' set conditional formatting for yearly change
                If yearlyChange > 0 Then
                    ws.Range("J" & summaryTableRow + 1).Interior.Color = vbGreen
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & summaryTableRow + 1).Interior.Color = vbRed
                Else
                    ws.Range("J" & summaryTableRow + 1).Interior.Color = vbYellow
                End If
                
                ' reset values for next ticker
                summaryTableRow = summaryTableRow + 1
                totalVolume = 0
                
            End If
            
        Next i
        
    Next ws
    
End Sub

