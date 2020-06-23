Attribute VB_Name = "Module1"
Sub TickerCounter1()
Dim WS As Worksheet
For Each WS In Worksheets
'2. Add Heading for Summary Table
        WS.Range("I1").Value = "Ticker Symbol"
        WS.Range("J1").Value = "Yearly Change"
        WS.Range("K1").Value = "% Change"
        WS.Range("L1").Value = "Total Stock Volume"
        WS.Range("O1").Value = "Greatest % Increase"
        WS.Range("O2").Value = "Greatest % Decrease"
        WS.Range("O3").Value = "Greatest Total Stock Volume"
'3. Dim Variables and their values if need be
        Dim OPrice As Double
        Dim CPrice As Double
        Dim TSymbol As String
        Dim YChange As Double
        Dim PChange As Double
        Dim TVolume As LongLong
            TVolume = 0
        Dim SummaryTable As Long
         SummaryTable = 2
        Dim GPIncrease As Double
            GPIncrease = 0
        Dim GPDecrease As Double
            GPDecrease = 0
        Dim GTVolume As Long
            GTVolume = 0
            'Determine Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Dim i As Long
'4. Run a loop inside of that loop that will run for each data table
    'Define how long loop will be for each Worksheet
        For i = 2 To LastRow
    'Determine which value it should pick up from Column A by determining if it is different than the one before it.
    'IF it is different then the following IF statement will run
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        'Determine The name of the Stock, and place the value in the "Ticker Symbol" in Row "I"
        TSymbol = WS.Cells(i, 1).Value
        WS.Range("I" & SummaryTable).Value = TSymbol
        'Determine The Yearly Change, and place the in the "Yearly Change" in Row "J"
         'Do this by first determining  open price and closed price first
            OPrice = WS.Cells(2, 3).Value
            CPrice = WS.Cells(i, 6).Value
            YChange = CPrice - OPrice
            WS.Range("J" & SummaryTable).Value = YChange
    
        'Determine The % Change, and place the in the "% Change" in Row "K"
         'Do this by doing (PChange / OPrice) * 100
            PChange = YChange / OPrice
            WS.Range("K" & SummaryTable).NumberFormat = "0.00%"
            WS.Range("K" & SummaryTable).Value = PChange
            
        'Determine The % Change, and place the in the "Total Stock Volume" in Row "L"
         'Do this by taking the number in the given cell and adding it  to row L
            TVolume = TVolume + Cells(i, 7).Value
            WS.Range("L" & SummaryTable).Value = TVolume
    
        
        'Go to the next row of summary Table
        SummaryTable = SummaryTable + 1
        'Reset Total Stock Volume to 0
        TVolume = 0
    
        Else
             'If Ticker Value in the next row is the same, Continue to add
        
        TVolume = TVolume + WS.Cells(i, 7).Value
    
        End If
    
        Next i
        
        'Formatting the YChange Cells
        YLastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
        For J = 2 To YLastRow
            'If The value in a given cell is >= to 0 then the Color Is Green
            If (WS.Cells(J, 10).Value >= 0) Then
                WS.Cells(J, 10).Interior.ColorIndex = 4
            'Else, the color of the cell will be red
            Else: WS.Cells(J, 10).Interior.ColorIndex = 3
        
            End If
        Next J
    

        
Next WS

End Sub


