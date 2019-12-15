Sub stonks()

    Dim tickerSymbol As String
    
    Dim yearlyChange As Double
    
    Dim percentChange As Double
    
    Dim totalStockVolume As Double

    Dim counter As Double

    Dim openingPrice As Double

    Dim closingPrice As Double

    Dim colorRed As Integer
    colorRed = 3

    Dim colorGreen As Integer
    colorGreen = 4
    
    totalStockVolume = 0
    
    Dim outputTableRow As Integer
    outputTableRow = 2
    
    Dim LastRow As Long
    LastRow = ActiveWorkbook.Worksheets("2014").Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    For i = 2 To LastRow
        
        
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
            
            counter = counter + 1
            
        Else
        
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
            
            openingPrice = ActiveWorkbook.Worksheets("2014").Cells(i - counter, 3).Value

            closingPrice = ActiveWorkbook.Worksheets("2014").Cells(i, 6).Value

            yearlyChange = closingPrice - openingPrice
            

            
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If

            tickerSymbol = Cells(i, 1).Value
            
            Cells(outputTableRow, 9).Value = tickerSymbol
            
            Cells(outputTableRow, 10).Value = yearlyChange
            If yearlyChange >= 0 Then
                Cells(outputTableRow, 10).Interior.ColorIndex = colorGreen

            Else
                Cells(outputTableRow, 10).Interior.ColorIndex = colorRed
            End If
            Cells(outputTableRow, 10).NumberFormat = "0.00"
            
            Cells(outputTableRow, 11).Value = percentChange
            
            Cells(outputTableRow, 11).NumberFormat = "0.00%"
            
            Cells(outputTableRow, 12).Value = totalStockVolume
            
            totalStockVolume = 0
            
            outputTableRow = outputTableRow + 1
            
        End If
        
    Next i
    
           

End Sub


