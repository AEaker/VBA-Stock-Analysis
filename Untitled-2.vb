Sub StockChangeCheck():
   
    Dim ws As Worksheet
    Dim Tick As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearChange As Double
    Dim PerChange As Double
    Dim TSV As Double
    
    
    Dim count As Integer
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2

    'Do I even need this?
    Dim MaxInc As Double
    Dim MaxIncName As String
    Dim MaxDec As Double
    Dim MaxDecName As String
    Dim MaxTotVolName As String
    Dim MaxTotVol As Double
    
   
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "% Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("M2").Value = "Greatest % Increase"
        ws.Range("M3").Value = "Greatest % Decrease"
        ws.Range("M4").Value = "Greatest Total Volume"
        ws.Range("N1").Value = "Ticker"
        ws.Range("O1").Value = "Value"
    
    TSV = 0
    YearClose = 0
    YearOpen = 0
    YearChange = 0
    PerChange = 0
    count = 0
    
    Tick_last_row = ws.Range("A2").End(xlDown).Row
    For r = 2 To Tick_last_row
    On Error Resume Next
        If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
            Tick = ws.Cells(r, 1).Value
            TSV = ws.Cells(r, 7).Value + TSV
            YearOpen = ws.Cells(r - count, 3).Value
            YearClose = ws.Cells(r, 6).Value
            YearChange = YearClose - YearOpen
            PerChange = YearChange / YearOpen
            ws.Range("I" & SummaryTableRow).Value = Tick
            ws.Range("J" & SummaryTableRow).Value = YearChange
            ws.Range("K" & SummaryTableRow).Value = Format(PerChange, "00.00%")
            ws.Range("L" & SummaryTableRow).Value = TSV
            SummaryTableRow = SummaryTableRow + 1
           
            TSV = 0
            Tick = 0
            YearOpen = 0
            YearClose = 0
            YearChange = 0
            PerChange = 0
            count = 0
        Else
            TSV = TSV + Cells(r, 7).Value
            count = count + 1
        End If
       Next r
        
        'Bonus
        ResultTickLast = ws.Range("I2").End(xlDown).Row
        
        MaxInc = ws.Application.WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("O2").Value = Format(MaxInc, "00.00%")
        MaxDec = ws.Application.WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("O3").Value = Format(MaxDec, "00.00%")
        MaxTotVol = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("O4").Value = MaxTotVol
            
        For r = 2 To ResultTickLast
        If ws.Cells(r, 12).Value = MaxTotVol Then
            MaxTotVolName = ws.Cells(r, 9).Value
            Else
            End If
        If ws.Cells(r, 11).Value = MaxInc Then
            MaxIncName = ws.Cells(r, 9).Value
            Else
            End If
        If ws.Cells(r, 11).Value = MaxDec Then
            MaxDecName = ws.Cells(r, 9).Value
            Else
            End If
            
        ws.Range("N4").Value = MaxTotVolName
        ws.Range("N3").Value = MaxDecName
        ws.Range("N2").Value = MaxIncName
        
        'Formatting
        If ws.Cells(r, 10).Value <= 0 Then
                ws.Cells(r, 10).Interior.Color = rgbRed
            Else
                ws.Cells(r, 10).Interior.Color = rgbGreen
            End If
        Next r
        SummaryTableRow = 2
       
        'Make everything look nice
        ws.Cells.EntireColumn.AutoFit
    Next ws
End Sub