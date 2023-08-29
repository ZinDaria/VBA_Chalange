# VBA_Chalange
Sub StockAnalysis()
    Dim i As Long
    Dim CurrentRow As String
    Dim NextRow As String
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim LastRow As Long
    Dim GroupNo As Long
    Dim CurrentSheet As Worksheet
    Dim YearlyChangeCell As Range
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim TickerGreatestIncrease As String
    Dim TickerGreatestDecrease As String
    Dim TickerGreatestTotalVolume As String
    Dim CountGreatestIncrease As Long
    Dim CountGreatestDecrease As Long
    Dim CountGreatestTotalVolume As Long
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    TickerGreatestIncrease = ""
    TickerGreatestDecrease = ""
    TickerGreatestTotalVolume = ""
    CountGreatestIncrease = 0
    CountGreatestDecrease = 0
    CountGreatestTotalVolume = 0
    
    For Each CurrentSheet In ActiveWorkbook.Worksheets
        TotalVolume = 0
        GroupNo = 1
        LastRow = CurrentSheet.Cells(CurrentSheet.Rows.Count, 1).End(xlUp).Row
        
        CurrentSheet.Cells(1, 9).Value = "Ticker"
        CurrentSheet.Cells(1, 10).Value = "Yearly Change"
        CurrentSheet.Cells(1, 11).Value = "Percent Change"
        CurrentSheet.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            CurrentRow = CurrentSheet.Cells(i, 1).Value
            NextRow = CurrentSheet.Cells(i + 1, 1).Value
            
            If NextRow <> CurrentRow Then
                ClosePrice = CurrentSheet.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                
                Set YearlyChangeCell = CurrentSheet.Cells(GroupNo + 1, 10)
                YearlyChangeCell.Value = YearlyChange
                
                If YearlyChange >= 0 Then
                    YearlyChangeCell.Interior.Color = RGB(0, 255, 0)
                Else
                    YearlyChangeCell.Interior.Color = RGB(255, 0, 0)
                End If
                
                CurrentSheet.Cells(GroupNo + 1, 9).Value = CurrentRow
                CurrentSheet.Cells(GroupNo + 1, 11).Value = PercentChange
                CurrentSheet.Cells(GroupNo + 1, 12).Value = TotalVolume
                
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    TickerGreatestIncrease = CurrentRow
                ElseIf PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    TickerGreatestDecrease = CurrentRow
                End If
                
                If TotalVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalVolume
                    TickerGreatestTotalVolume = CurrentRow
                End If
                
                TotalVolume = 0
                OpenPrice = CurrentSheet.Cells(i + 1, 3).Value
                GroupNo = GroupNo + 1
            Else
                TotalVolume = TotalVolume + CurrentSheet.Cells(i, 7).Value
                If i = 2 Then
                    OpenPrice = CurrentSheet.Cells(i, 3).Value
                End If
            End If
            
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                TickerGreatestIncrease = CurrentRow
                CountGreatestIncrease = 1
            ElseIf PercentChange = GreatestIncrease And TickerGreatestIncrease <> "" Then
                CountGreatestIncrease = CountGreatestIncrease + 1
            End If
            
            If PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                TickerGreatestDecrease = CurrentRow
                CountGreatestDecrease = 1
            ElseIf PercentChange = GreatestDecrease And TickerGreatestDecrease <> "" Then
                CountGreatestDecrease = CountGreatestDecrease + 1
            End If
            
            If TotalVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalVolume
                TickerGreatestTotalVolume = CurrentRow
                CountGreatestTotalVolume = 1
            ElseIf TotalVolume = GreatestTotalVolume And TickerGreatestTotalVolume <> "" Then
                CountGreatestTotalVolume = CountGreatestTotalVolume + 1
            End If
        Next i
        
        CurrentSheet.Range(CurrentSheet.Cells(2, 11), CurrentSheet.Cells(LastRow, 11)).NumberFormat = "0.00%"
        
        CurrentSheet.Range("N2").Value = "Greatest % Increase"
        CurrentSheet.Range("N3").Value = "Greatest % Decrease"
        CurrentSheet.Range("N4").Value = "Greatest Total Value"
        CurrentSheet.Range("O1").Value = "Ticker"
        CurrentSheet.Range("P1").Value = "Value"
        CurrentSheet.Range("O2").Value = TickerGreatestIncrease
        CurrentSheet.Range("O3").Value = TickerGreatestDecrease
        CurrentSheet.Range("O4").Value = TickerGreatestTotalVolume
        
        If TickerGreatestIncrease <> "" Then
            CurrentSheet.Range("P2").Offset(1, 0).Value = GreatestIncrease
            CurrentSheet.Range("P2").Offset(1, 0).NumberFormat = "0.00%"
        End If
        If TickerGreatestDecrease <> "" Then
            CurrentSheet.Range("P3").Offset(1, 0).Value = GreatestDecrease
            CurrentSheet.Range("P3").Offset(1, 0).NumberFormat = "0.00%"
        End If
        If TickerGreatestTotalVolume <> "" Then
            CurrentSheet.Range("P4").Offset(1, 0).Value = GreatestTotalVolume
        End If
    Next CurrentSheet
End Sub
