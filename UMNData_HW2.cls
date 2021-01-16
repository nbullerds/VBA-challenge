VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockAnalysis()

    Dim openingValue As Double
    Dim closingValue As Double
    Dim totalVolume As Double
    Dim countOfTicker As Double
    Dim numberOfTickers As Double
    
    ' Iterate through every worksheet
    For Each ws In Worksheets
        
        'Set up output columns and rows on each sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Changed"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' set variables to initial values at the beginning of each sheet in ws
        numberOfTickers = 0
        openingValue = 0
        closingValue = 0
        totalVolume = 0
        countOfTicker = 0

        ' Iterate through every row in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For x = 2 To LastRow

            ' Check if the current row has the same ticker symbol as the next
            If (ws.Cells(x, 1).Value = ws.Cells(x + 1, 1).Value) Then

                totalVolume = totalVolume + ws.Cells(x, 7).Value
                countOfTicker = countOfTicker + 1

            Else
            
                ' List the ticker value for this ticker
                ws.Cells(numberOfTickers + 2, 9).Value = ws.Cells(x, 1).Value
                
                ' Calculate the openning and closing value of the ticker for this year
                closingValue = ws.Cells(x, 6).Value
                openingValue = ws.Cells(x - (countOfTicker), 3).Value
                countOfTicker = 0
                
                ' Calculate yearly change and highlight accordingly
                ws.Cells(numberOfTickers + 2, 10).Value = (closingValue - openingValue)
                If ws.Cells(numberOfTickers + 2, 10).Value < 0 Then
                    ws.Cells(numberOfTickers + 2, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(numberOfTickers + 2, 10).Interior.ColorIndex = 4
                End If
                
                ' Calculate yearly percent change
                If openingValue <> 0 Then
                    ws.Cells(numberOfTickers + 2, 11).Value = FormatPercent(((closingValue - openingValue) / openingValue), 2)
                Else
                    ws.Cells(numberOfTickers + 2, 11).Value = "n/a"
                End If
                
                ' Calculate totalVolume
                ws.Cells(numberOfTickers + 2, 12).Value = (totalVolume + ws.Cells(x, 7).Value)
                totalVolume = 0
                
                numberOfTickers = numberOfTickers + 1

            End If

        Next
        
    Next
    
End Sub
