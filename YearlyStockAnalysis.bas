Attribute VB_Name = "Module1"
Sub YearlyStockAnalysis()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRowIndex As Long
    
   
    Set ws = ThisWorkbook.Worksheets("2018")
    
   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Naming Summary Table Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    summaryTableRowIndex = 2
    totalVolume = 0
    
    ' Looping through the Stock data
    For i = 2 To lastRow
    
        ' Check if the ticker symbol has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' Store the closing price
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculate the yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate the percent change
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
                
            Else
                percentChange = 0
                
            End If
            
            ' Print the values to the summary table
            ws.Cells(summaryTableRowIndex, 9).Value = ticker
            ws.Cells(summaryTableRowIndex, 10).Value = yearlyChange
            ws.Cells(summaryTableRowIndex, 11).Value = percentChange
            ws.Cells(summaryTableRowIndex, 12).Value = totalVolume
            
            ' Format the percent change as a percentage
            ws.Cells(summaryTableRowIndex, 11).NumberFormat = "0.00%"
            
            ' Reset variables for the next ticker symbol
            summaryTableRowIndex = summaryTableRowIndex + 1
            
            totalVolume = 0
        End If
        
        ' Adding the stock volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        ' Store the opening price
        If openingPrice = 0 Then
            openingPrice = ws.Cells(i, 3).Value
        End If
                        
    Next i
    
    ' Auto-fit columns in the summary table
    ' ws.Columns("I:L").AutoFit
       
End Sub

