VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub analysis()
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    TotalVolume = 0
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Dim OpenPrice As Double
    OpenPrice = Range("C2").Value
    
    Dim ClosePrice As Double
    Dim LastRow As Long
    Dim TickerRow As Integer
    TickerRow = 2
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
    Dim i As Long
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Range("I" & TickerRow).Value = Ticker
            Range("L" & TickerRow).Value = TotalVolume
            ClosePrice = Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
            Range("J" & TickerRow).Value = YearlyChange
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
            Range("K" & TickerRow).Value = PercentChange
            Range("K" & TickerRow).NumberFormat = "0.00%"

            TickerRow = TickerRow + 1
            
            TotalVolume = 0
            
            OpenPrice = Cells(i + 1, 3)
        Else
            TotalVolume = TotalVolume + Cells(i, 7).Value
        End If
    Next i
    
    SummaryLastRow = Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To SummaryLastRow
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    For i = 2 To SummaryLastRow
        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & SummaryLastRow)) Then
            Range("O2").Value = Cells(i, 9).Value
            Range("P2").Value = Cells(i, 11).Value
            Range("P2").NumberFormat = "0.00%"
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & SummaryLastRow)) Then
            Range("O3").Value = Cells(i, 9).Value
            Range("P3").Value = Cells(i, 11).Value
            Range("P3").NumberFormat = "0.00%"
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & SummaryLastRow)) Then
            Range("O4").Value = Cells(i, 9).Value
            Range("P4").Value = Cells(i, 12).Value
        End If
    Next i
            
End Sub
