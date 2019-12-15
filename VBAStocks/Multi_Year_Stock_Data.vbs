Attribute VB_Name = "Module1"
Sub StockAnalysisSummary()

For Each ws In Worksheets

' Header row for Stock Analysis
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ws.Rows(1).Font.Bold = True
    ws.Columns("O").Font.Bold = True
    
' Declare Variables
    Dim Ticker As String
    Dim LastRow As Long
    Dim TtlTickerVol As Double
    Dim SummaryTblRow As Long
    Dim BeginPrice As Double
    Dim ClosePrice As Double
    Dim PriceChange As Double
    Dim PrevAmt As Long
    Dim PercentChange As Double
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim LastRowValue As Long
    Dim GreatestTtlVol As Double

' Initialize Variables
    GreatestInc = 0
    PrevAmt = 2
    SummaryTblRow = 2
    TtlTickerVol = 0
    GreatestDec = 0
    GreatestTtlVol = 0

' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row
    For i = 2 To LastRow
    
        ' Add Total Volume
        TtlTickerVol = TtlTickerVol + ws.Cells(i, 7).Value
        
        ' Check for same Ticker If Not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryTblRow).Value = Ticker
            ws.Range("L" & SummaryTblRow).Value = TtlTickerVol
            TtlTickerVol = 0
            
            BeginPrice = ws.Range("C" & PrevAmt)
            ClosePrice = ws.Range("F" & i)
            PriceChange = ClosePrice - BeginPrice
            ws.Range("J" & SummaryTblRow).Value = PriceChange

        ' Calculate Percent Change
            If BeginPrice = 0 Then
                PercentChange = 0
            Else
                BeginPrice = ws.Range("C" & PrevAmt)
                PercentChange = PriceChange / BeginPrice
            End If
        
        ' Format PercentChange as 0.00%
            ws.Range("K" & SummaryTblRow).NumberFormat = "0.00%"
            ws.Range("K" & SummaryTblRow).Value = PercentChange

        ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
            If ws.Range("J" & SummaryTblRow).Value >= 0 Then
                ws.Range("J" & SummaryTblRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & SummaryTblRow).Interior.ColorIndex = 3
            End If
            
        ' Incremeant Row Count
            SummaryTblRow = SummaryTblRow + 1
            PrevAmt = i + 1
        
        End If
    Next i
        
    ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
    ' Start Loop For Final Results
    For i = 2 To LastRow
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
        End If
    
    Next i
    
    ' Format Double To Include % Symbol And Two Decimal Places
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
            
    ' Format Table Columns To Auto Fit
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub
