Sub hw()
'set up worksheet loop
    For Each ws In Worksheets
        'label columns to put work in
        ws.Range("M1").Value = "ticker smbl"
        ws.Range("N1").Value = "yr change"
        ws.Range("O1").Value = "yr % change"
        ws.Range("P1").Value = "ttl volume"
        
        'set up variables for data table
        Dim Ticker As String
        Dim TotalVolume As Double
        TotalVolume = 0
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YrChange As Double
        Dim PChange As Double
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        
        'determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            'loop!
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'get ticker for summary table
                Ticker = ws.Cells(i, 1).Value
                'get open and close price to calculate changes
                OpenPrice = ws.Cells(i, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                'calculate changes
                YrChange = ClosePrice - OpenPrice
                PChange = (YrChange / ClosePrice) * 100
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                'Write to summary table
                ws.Range("M" & Summary_Table_Row).Value = Ticker
                ws.Range("N" & Summary_Table_Row).Value = YrChange
                ws.Range("O" & Summary_Table_Row).Value = PChange
                ws.Range("P" & Summary_Table_Row).Value = TotalVolume
                Summary_Table_Row = Summary_Table_Row + 1
                'Reset volume
                TotalVolume = 0
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub
