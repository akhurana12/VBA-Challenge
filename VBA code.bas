Attribute VB_Name = "Module1"
Sub YearlyStockData()

    'Describe all variables
    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalStockVolume As Double
    Dim GreatestPercentageIncrease As Double
    Dim GreatestPercentageDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim OpeningPrice As Double
    Dim TickerRow As Integer
    Dim WorksheetName As String
    
    For Each ws In Worksheets
    
        
        Dim PerChange As Double
        Dim TickCount As Long
        Dim j As Long
        Dim LastRow As Long
        Dim LastRowI As Long
        
        ' Get the WorksheetName
        WorksheetName = ws.Name
        
        ' Initialize the variables
        GreatestPercentageIncrease = 0
        GreatestPercentageDecrease = 0
        GreatestTotalVolume = 0
        TickCount = 2
        
        ' Generate Headers for all the rows and columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Set start row to 2
        j = 2
        
        ' Determine the last row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows
        For i = 2 To LastRow
        
            ' Check if ticker name changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerSymbol = ws.Cells(i, 1).Value
                ws.Cells(TickCount, 9).Value = TickerSymbol
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                ' Conditional formatting
                If ws.Cells(TickCount, 10).Value < 0 Then

                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                Else

                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                End If
                
                ' Calculate percent change for column K
                If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ' Percent formatting
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                Else
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                End If
                
                ' Calculate total volume for column L
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                ' Increase TickCount by 1
                TickCount = TickCount + 1
                
                ' Set new start row of the ticker set
                j = i + 1
            End If
        Next i
        
        ' Find last non-blank cell in column I
        LastRowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        ' Prepare for summary
        GreatestTotalVolume = ws.Cells(2, 12).Value
        GreatestPercentageIncrease = ws.Cells(2, 11).Value
        GreatestPercentageDecrease = ws.Cells(2, 11).Value
        
        ' Loop for summary
        For i = 2 To LastRowI
        
            ' For greatest total volume--check if the next value is larger--if yes take over a new value and populate ws.Cells
            If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
            
            ' For greatest increase--check if the next value is larger--if yes take over a new value and populate ws.Cells
            If ws.Cells(i, 11).Value > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            
            ' For greatest decrease--check if the next value is smaller--if yes take over a new value and populate ws.Cells
            If ws.Cells(i, 11).Value < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i,

