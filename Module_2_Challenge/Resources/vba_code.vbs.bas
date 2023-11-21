Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
    
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables for the first row of each sheet
        ticker = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        summaryRow = 2 '
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        maxIncreaseTicker = ""
        maxDecreaseTicker = ""
        maxVolumeTicker = ""
        

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        For Row = 2 To lastRow
        
            ' Check if the ticker symbol has changed
            If ws.Cells(Row, 1).Value <> ticker Then
            
                ' Calculate yearly change and percent change
                closingPrice = ws.Cells(Row - 1, 6).Value
                yearlyChange = closingPrice - openingPrice
                
                If openingPrice <> 0 Then
                    percentChange = ((ws.Cells(Row - 1, 6).Value / openingPrice) - 1)
                Else
                    percentChange = 0
                End If
                
                ' Output data to the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
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
                

                summaryRow = summaryRow + 1
                

                ticker = ws.Cells(Row, 1).Value
                openingPrice = ws.Cells(Row, 3).Value
                totalVolume = 0
                
            End If
            
            ' Accumulate total volume for the ticker
            totalVolume = totalVolume + ws.Cells(Row, 7).Value
            
        Next Row
        
        closingPrice = ws.Cells(lastRow, 6).Value
        yearlyChange = closingPrice - openingPrice
        
        If openingPrice <> 0 Then
            percentChange = ((closingPrice / openingPrice) - 1) * 100
        Else
            percentChange = 0
        End If
        
        ' Output data to the summary table
        ws.Cells(summaryRow, 9).Value = ticker
        ws.Cells(summaryRow, 10).Value = yearlyChange
        ws.Cells(summaryRow, 11).Value = percentChange
        ws.Cells(summaryRow, 12).Value = totalVolume
  
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = maxIncrease
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = maxDecrease
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(4, 16).Value = maxVolume
        
        ApplyConditionalFormatting ws
    
    Next ws

End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)
    ' Apply conditional formatting to Percentage Change
    With ws.Range(ws.Cells(2, 11), ws.Cells(ws.Cells(ws.Rows.Count, "K").End(xlUp).Row, 11)).FormatConditions
        .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .Item(1).Interior.Color = RGB(0, 255, 0)
        
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(2).Interior.Color = RGB(255, 0, 0)
    End With
    
    ' Apply conditional formatting to Yearly Change
    With ws.Range(ws.Cells(2, 10), ws.Cells(ws.Cells(ws.Rows.Count, "J").End(xlUp).Row, 10)).FormatConditions
        .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .Item(1).Interior.Color = RGB(0, 255, 0)
        
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(2).Interior.Color = RGB(255, 0, 0)
    End With

    ws.Range(ws.Cells(2, 11), ws.Cells(ws.Cells(ws.Rows.Count, "K").End(xlUp).Row, 11)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(2, 16), ws.Cells(3, 16)).NumberFormat = "0.00%"
    
End Sub


