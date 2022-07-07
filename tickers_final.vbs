Attribute VB_Name = "Module1"
Sub stocks()
    
    'Find the number of rows to loop over
    Dim rowcount As Double
    rowcount = Range("A1").End(xlDown).Row - 1
    'MsgBox (rowcount)
    
        
    'Declare variables for stock data
    Dim tickers() As String
    Dim prices() As Double
    Dim volumes() As Double
    Dim j As Integer
    
    'Count the number of unique tickers
    Dim tickercount As Integer
    tickercount = 0
    For i = 1 To rowcount
        If Cells(i + 2, 1).Value <> Cells(i + 1, 1).Value Then
            tickercount = tickercount + 1
        Else
        End If
    Next i
    'MsgBox (tickercount)
    
    
    'Resize dynamic arrays
    ReDim tickers(tickercount - 1)
    ReDim prices(tickercount - 1, 1)
    ReDim volumes(tickercount - 1)
    
    'Intialize variables
    tickers(0) = Range("A2").Value
    prices(0, 0) = Range("C2").Value
    volumes(0) = 0
    j = 0
   
    'Loop through the stock data
    For i = 1 To rowcount
        If Cells(i + 2, 1) = Cells(i + 1, 1) Then
            volumes(j) = volumes(j) + Cells(i + 1, 7).Value
        Else
            volumes(j) = volumes(j) + Cells(i + 1, 7).Value
            prices(j, 1) = Cells(i + 1, 6).Value
            
            If j < (tickercount - 1) Then
                'Step to next j and reintitalize data
                j = j + 1
                tickers(j) = Cells(i + 2, 1).Value
                prices(j, 0) = Cells(i + 2, 3).Value
                volumes(j) = 0
            Else
            End If
        End If
    Next i
    
    'Output the results
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    For i = 1 To tickercount
        Cells(i + 1, 9).Value = tickers(i - 1)
        Cells(i + 1, 10).Value = prices(i - 1, 1) - prices(i - 1, 0)
        'Apply conditional formatting
        If Cells(i + 1, 10).Value > 0 Then
            Cells(i + 1, 10).Interior.ColorIndex = 4
        Else
            Cells(i + 1, 10).Interior.ColorIndex = 3
        End If
        Cells(i + 1, 11).Value = (prices(i - 1, 1) - prices(i - 1, 0)) / prices(i - 1, 0)
        Cells(i + 1, 12).Value = volumes(i - 1)
    Next i
    
    'Apply formatting to percent change column
    
    For Each cell In Range(Cells(2, 11), Cells(tickercount + 1, 11))
        cell.NumberFormat = "0.00%"
    Next
    
End Sub
