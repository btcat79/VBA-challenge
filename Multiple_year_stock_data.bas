Attribute VB_Name = "Module3"
Sub AnalyzeStockData()

    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim totalVolume As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim summaryRow As Long
    Dim startRow As Long
    
    ' Array of sheet names to process
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Loop through each sheet
    For Each ws In ThisWorkbook.Sheets(sheetNames)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables for tracking the greatest changes
        maxPercentIncrease = -1
        maxPercentDecrease = 1
        maxVolume = 0
        summaryRow = 2
        
        ' Loop through all rows to process each stock
        i = 2 ' Start from the first row of data
        
        Do While i <= lastRow
            ticker = ws.Cells(i, "A").Value
            startRow = i
            
            ' Find the last row for this ticker
            Do While i <= lastRow And ws.Cells(i, "A").Value = ticker
                i = i + 1
            Loop
            
            ' Calculate the quarterly change
            openPrice = ws.Cells(startRow, "C").Value ' First day's opening price
            closePrice = ws.Cells(i - 1, "F").Value ' Last day's closing price
            quarterlyChange = closePrice - openPrice
            
            ' Sum total volume over the quarter
            totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, "G"), ws.Cells(i - 1, "G")))
            
            ' Calculate the percentage change
            If openPrice <> 0 Then
                percentChange = quarterlyChange / openPrice
            Else
                percentChange = 0
            End If
            
            ' Add data to the summary table
            ws.Cells(summaryRow, "I").Value = ticker
            ws.Cells(summaryRow, "J").Value = quarterlyChange
            ws.Cells(summaryRow, "K").Value = percentChange
            ws.Cells(summaryRow, "L").Value = totalVolume
            
            ' Format column K as percent with two decimal places
            ws.Cells(summaryRow, "K").NumberFormat = "0.00%"
            
            ' Determine the greatest increase, decrease, and max volume
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxIncreaseTicker = ticker
            End If
            
            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxDecreaseTicker = ticker
            End If
            
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolumeTicker = ticker
            End If
            
            summaryRow = summaryRow + 1
        Loop
        
        ' Output the greatest increase, decrease, and max volume
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(2, "P").Value = maxIncreaseTicker
        ws.Cells(2, "Q").Value = maxPercentIncrease
        
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(3, "P").Value = maxDecreaseTicker
        ws.Cells(3, "Q").Value = maxPercentDecrease
        
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(4, "P").Value = maxVolumeTicker
        ws.Cells(4, "Q").Value = maxVolume
        
        ' Format cells Q2 and Q3 as percent with two decimal places
        ws.Cells(2, "Q").NumberFormat = "0.00%"
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        
        ' Add headers for the greatest values table
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
    Next ws
    
  
End Sub


