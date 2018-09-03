Sub StockData()

    'Defining variables
    Dim ws as Worksheet
    Dim StockVolume As Double
    Dim YearlyChange as Double
    Dim TickerStart as Double
    Dim TickerEnd as Double
    Dim LRow As Double

    'Setting loops
    For Each ws in WorkSheets
        ws.activate
        LRow = Range("A" & Rows.Count).End(xlUp).Row
        j = 1
        For i = 2 To LRow

            'Conditional for defining start of a stock ticker and creating new row of output
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                j = j + 1
                Cells(j, 9).Value = Cells(i, 1).Value
                TickerStart = i

            'Conditoinal for defining end of a stock ticker and summing the range between beginning and end
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                TickerEnd = i
                StockVolume = WorksheetFunction.Sum(Range("G" & TickerStart & ":G" & TickerEnd))
                YearlyChange = Range("F" & TickerEnd).Value - Range("C" & TickerStart).Value

                'reporting and formating yearly change'
                Cells(j, 10).Value = YearlyChange
                If Cells(j, 10).Value > 0 then
                  Cells(j, 10).interior.colorindex = 4
                ElseIf Cells(j, 10).Value < 0 then
                  Cells(j, 10).interior.colorindex = 3
                Else
                  Cells(j, 10).interior.colorindex = 0
                End If

                'reporting percent change; conditional to avoid stack overflow from any zero values
                If Range("C" & TickerStart).Value <> 0 Then
                  Cells(j, 11).Value = YearlyChange / Range("C" & TickerStart).Value
                Else
                  Cells(j, 11).Value = 0
                End If
                Cells(j, 11).NumberFormat = "0.0%"

                'reporting total stock volume'
                Cells(j, 12).Value = StockVolume
            End If
        Next i

        'defining Max % increase, % decrease, and total volume'
        Dim MaxIncrease as Double
        Dim MaxDecrease as Double
        Dim MaxTotalVolume as Double

        'Max increase'
        MaxIncrease = 0
        For i = 2 to 5000
          If Cells(i, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(i, 11).Value
            MaxIncreaseTicker = Cells(i, 9).Value
          End If
        Next i
        Cells(2, 16).Value = MaxIncreaseTicker
        Cells(2, 17).Value = MaxIncrease
        Cells(2, 17).NumberFormat = "0.0%"

        'Max decrease'
        MaxDecrease = 0
        For i = 2 to 5000
          If Cells(i, 11).Value < MaxDecrease Then
            MaxDecrease = Cells(i, 11).Value
            MaxDecreaseTicker = Cells(i, 9).Value
          End If
        Next i
        Cells(3, 16).Value = MaxDecreaseTicker
        Cells(3, 17).Value = MaxDecrease
        Cells(3, 17).NumberFormat = "0.0%"

        'Max total volume'
        MaxTotalVolume = 0
        For i = 2 to 5000
          If Cells(i, 12).Value > MaxTotalVolume Then
            MaxTotalVolume = Cells(i, 12).Value
            MaxTotalVolumeTicker = Cells(i, 9).Value
          End If
        Next i
        Cells(4, 16).Value = MaxTotalVolumeTicker
        Cells(4, 17).Value = MaxTotalVolume

    Next ws
End Sub
