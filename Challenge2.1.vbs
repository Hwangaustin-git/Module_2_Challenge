Sub Analysis()

    ' Declaring variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Set initial values for tracking the maximums
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0

    ' Go through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' summary strats on row 2
        summaryRow = 2
        totalVolume = 0
        ' first open price
        openPrice = ws.Cells(2, 3).Value

        ' Adding fixed headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Loop through the stock data
        For i = 2 To lastRow
            
            ' Add up the total volume for this ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if it got to the last row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                
                ' change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice)
                Else
                    percentChange = 0
                End If

                ' Write the results in column I to L
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryRow, 12).Value = totalVolume

                ' Highlighting quarterly change column
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed
                End If

                ' Update max values
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If

                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If

                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                ' Move to the next row
                summaryRow = summaryRow + 1

                ' Reset for the next ticker
                totalVolume = 0
                If i < lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                End If

            End If

        Next i

        ' Write the maximum values to columns O~Q
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = maxDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume

    Next ws

End Sub


