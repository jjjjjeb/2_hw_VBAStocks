' This script will:
    ' - Loop through all the stocks for one year
    '   for each run and take the following information.
        ' 1 - The ticker symbol.
        ' 2 - Yearly change from opening price at the beginning of the year
        '     to the closing price at the end of the year.
        ' 3 - The percent change from opening price at the beginning of the year
        '     to the closing price at the end of the year.
        ' 4 - The total stock volume of the stock.
        ' 5 - Sets conditional formatting  that will highlight positive change 
        '     in green and negative change in red.
        ' 6 - Adds another table that states the greatest increase, decrease,
        '     and highest total volume tickers per worksheet.
'____________________________________________________________________________________________

Sub ABCtest()

' loop though all the sheets
Dim ws As Worksheet
For Each ws In Worksheets
      
' set variables and assign values for Ticker Symbol, Opening Price,
    ' Closing Price, Yearly Change, Total Stock Volume
    
    Dim TickerSymbol As String
    Dim TickerTotal As Integer
    Dim TicYearOpen As Double
    Dim TicYearClose As Double
    Dim TicYearChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
        StockVolume = 0
    Dim LastRow As Long
    
    ' Determine the Last Row
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Summary Table Headers
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     
    'Final table
    ws.Range("P1").Value = "Ticker Symbol"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
            
    'Keep track of the location for each stock ticker in  the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Keep track of previous amount
    Dim PreviousAmount As Long
    PreviousAmount = 2
    
    'Loop through tickers
    For I = 2 To LastRow
     
        'Add to total stock volume
        StockVolume = StockVolume + ws.Cells(I, 7).Value
     
        'Check if still w/in the same stock, if not...
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

'------TICKER & VOLUME

            ' Set ticker name
            TickerSymbol = ws.Cells(I, 1).Value
        
            'Print ticker in summ table
            ws.Range("I" & Summary_Table_Row).Value = TickerSymbol
        
            'Add to volume
            ws.Range("L" & Summary_Table_Row).Value = StockVolume
        
            'Reset
            StockVolume = 0
        
'------YEAR ANALYSIS
            
            'Set Year Open, Close & Change Names
            TicYearOpen = ws.Range("C" & PreviousAmount)
            TicYearClose = ws.Range("F" & I)
            TicYearChange = TicYearClose - TicYearOpen
            
            'Print Year Difference
            ws.Range("J" & Summary_Table_Row).Value = TicYearChange

'------GET % CHANGE
                
            If TicYearOpen = 0 Then
            PercentChange = 0
            
            Else
            TicYearOpen = ws.Range("C" & PreviousAmount)
            PercentChange = TicYearChange / TicYearOpen
            
            End If
                
            'Print % change ad make it look like %
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            
            'Add conditional formating for pos and neg values \(*.*)/
            If ws.Range("K" & Summary_Table_Row).Value < 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 54
            
            Else
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 43
            
            End If
                        
            ' Add one to the summary table row & previous amount values
            Summary_Table_Row = Summary_Table_Row + 1
            PreviousAmount = I + 1
      
      End If
      
      Next I

'------ HIGHEST % HIGHS, LOWEST % LOWS, GREATEST TOTAL VOL

            
            Dim maxIn As Long
            Dim maxDe As Long
            Dim maxVol As Double
            maxIn = 0
            maxDe = 0
            maxVol = 0
            Dim lr As Long
            
            'Get last row for final table
            lr = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
            
            'Start loop for final table & set conditions
            For j = 2 To lr
            
            If ws.Range("K" & j).Value > ws.Range("Q2").Value Then
            maxIn = ws.Range("K" & j).Value
            ws.Range("Q2").Value = maxIn
            ws.Range("P2").Value = ws.Range("I" & j).Value
            
            End If

            If ws.Range("K" & j).Value < ws.Range("Q3").Value Then
            maxDe = ws.Range("K" & j).Value
            ws.Range("Q3").Value = maxDe
            ws.Range("P3").Value = ws.Range("I" & j).Value
                
            End If

            If ws.Range("L" & j).Value > ws.Range("Q4").Value Then
            maxVol = ws.Range("L" & j).Value
            ws.Range("Q4").Value = maxVol
            ws.Range("P4").Value = ws.Range("I" & j).Value
            
            End If

            Next j
        
            'format %'s
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
            'format Table Columns To Auto Fit
            ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub

