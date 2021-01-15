Attribute VB_Name = "Module1"
Sub stocks()

'loop through each worksheet
For Each ws In Worksheets
    ws.Activate

    'define headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    'define variables
    Dim ticker As String
    Dim ticker_summary As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim volume As Double
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set variables
    ticker = ""
    ticker_summary = 0
    open_price = 0
    year_change = 0
    percent_change = 0
    volume = 0
    
    'loop through data
    For i = 2 To last_row
        ticker = Cells(i, 1).Value
        
        If open_price = 0 Then
            open_price = Cells(i, 3).Value
        End If
        
        volume = volume + Cells(i, 7).Value
        
        If ticker <> Cells(i + 1, 1).Value Then
            'ticker
            ticker_summary = ticker_summary + 1
            Cells(ticker_summary + 1, 9).Value = ticker
            
            'close price
            close_price = Cells(i, 6).Value
            
            'year change
            year_change = close_price - open_price
            
            Cells(ticker_summary + 1, 10).Value = year_change
        
            'color code
            If year_change < 0 Then
                Cells(ticker_summary + 1, 10).Interior.ColorIndex = 3
            ElseIf year_change > 0 Then
                Cells(ticker_summary + 1, 10).Interior.ColorIndex = 4
            End If
            
            'percent change
            If open_price = 0 Then
                percent_change = 0
            Else: percent_change = (year_change / open_price)
            End If
            
            Cells(ticker_summary + 1, 11).Value = Format(percent_change, "Percent")
            
            open_price = 0
            
            Cells(ticker_summary + 1, 12).Value = volume
                
            volume = 0
    
        End If
        
    Next i

Next ws

End Sub

