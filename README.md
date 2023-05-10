# VBA-challenge

Sub StockAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Integer

    For Each ws In ThisWorkbook.Worksheets
        ' Initialize summary row
        SummaryRow = 2
        
        ' Add header for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Find the last row with data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initialize YearlyOpen value
        YearlyOpen = ws.Cells(2, 3).Value

        ' Loop through all rows with data
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get the values for the calculations
                Ticker = ws.Cells(i, 1).Value
                YearlyClose = ws.Cells(i, 6).Value
                YearlyChange = YearlyClose - YearlyOpen
                If YearlyOpen <> 0 Then
                    PercentChange = (YearlyChange / YearlyOpen) * 100
                Else
                    PercentChange = 0#
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ' Write the summary to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange & "%"
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Reset variables for next ticker
                YearlyOpen = ws.Cells(i + 1, 3).Value
                TotalVolume = 0
                SummaryRow = SummaryRow + 1
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Format the summary table
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "0"
    Next ws
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize summary row
        SummaryRow = 2
        
        ' Add header for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Add header for greatest values table
        ws.Cells(1, 15).Value = "Category"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Find the last row with data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initialize YearlyOpen value
        YearlyOpen = ws.Cells(2, 3).Value

        ' Initialize variables for greatest values
        MaxPercentIncrease = 0
        MaxPercentDecrease = 0
        MaxTotalVolume = 0

        ' Loop through all rows with data
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get the values for the calculations
                Ticker = ws.Cells(i, 1).Value
                YearlyClose = ws.Cells(i, 6).Value
                YearlyChange = YearlyClose - YearlyOpen
                If YearlyOpen <> 0 Then
                    PercentChange = (YearlyChange / YearlyOpen) * 100
                Else
                    PercentChange = 0
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ' Write the summary to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange / 100
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Check for greatest values
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    MaxPercentIncreaseTicker = Ticker
                End If
                
                If PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    MaxPercentDecreaseTicker = Ticker
                End If
                
                If TotalVolume > MaxTotalVolume Then
                    MaxTotalVolume = TotalVolume
                                    MaxTotalVolumeTicker = Ticker
            End If

            ' Reset variables for next ticker
            YearlyOpen = ws.Cells(i + 1, 3).Value
            TotalVolume = 0
            SummaryRow = SummaryRow + 1
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
    Next i

    ' Write the greatest values to the table
    ws.Cells(2, 16).Value = MaxPercentIncreaseTicker
    ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
    ws.Cells(4, 16).Value = MaxTotalVolumeTicker
    ws.Cells(2, 17).Value = MaxPercentIncrease / 100
    ws.Cells(3, 17).Value = MaxPercentDecrease / 100
    ws.Cells(4, 17).Value = MaxTotalVolume

    ' Format the summary table and greatest values table
    ws.Columns("J").NumberFormat = "0.00"
    ws.Columns("L").NumberFormat = "0"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0.00E+00"
    ' Apply conditional formatting to the "Yearly Change" column
        With ws.Range("J2:J" & SummaryRow - 1)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1)
                .Interior.Color = RGB(0, 255, 0)
                .StopIfTrue = False
            End With

            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1)
                .Interior.Color = RGB(255, 0, 0)
                .StopIfTrue = False
            End With
        End With
        
    Next ws

End Sub
