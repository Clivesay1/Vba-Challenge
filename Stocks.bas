Attribute VB_Name = "Stocks"
Sub Stocks()

    Dim ticker As String
    Dim opening As Double, closing As Double
    Dim quarterly As Double, Percent_Change As Double, total_stock As Double
    Dim previous_price As Long, ticker_row As Long
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    total_stock = 0
    previous_price = 2
    ticker_row = 2
    
    end_ticker = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    For a = 2 To end_ticker:
        total_stock = total_stock + ws.Cells(a, 7).Value
        
        If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
            ticker = ws.Cells(a, 1).Value
            ws.Range("I" & ticker_row).Value = ticker
            ws.Range("L" & ticker_row).Value = total_stock
            total_stock = 0 'friend michael told me to reset here
            
            opening = ws.Range("C" & previous_price)
            closing = ws.Range("F" & a)
            quarterly = closing - opening
            ws.Range("J" & ticker_row).Value = quarterly
            
            If opening = 0 Then
                Percent_Change = 0
                Else
                    Percent_Change = quarterly / opening
            End If
            
            If ws.Range("J" & ticker_row).Value >= 0 Then
                ws.Range("J" & ticker_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & ticker_row).Interior.ColorIndex = 3
            End If
        
                ws.Range("K" & ticker_row).Value = Percent_Change
                ws.Range("K" & ticker_row).NumberFormat = "0.00%" 'per michael
            
            previous_price = a + 1
            ticker_row = ticker_row + 1
        End If
    Next a
            
            
    Dim greatest_increase As Double, greatest_decrease As Double, greatest_volume As Double
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Incease"
    ws.Range("O3").Value = "Greatest % Deacrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    percent_end = ws.Cells(Rows.Count, 11).End(xlUp).row
    
    For a = 2 To percent_end:
        If ws.Range("K" & a).Value > Cells(2, 17) Then
            greatest_increase = ws.Range("K" & a).Value
            ws.Range("R2").Value = greatest_increase
            ws.Range("Q2").Value = ws.Range("I" & a).Value
        End If
        
        If ws.Range("K" & a).Value < Cells(3, 17) Then
            greatest_decrease = ws.Range("K" & a).Value
            ws.Range("R3").Value = greatest_decrease
            ws.Range("Q3").Value = ws.Range("I" & a).Value
        End If
        
        If ws.Range("L" & a).Value > Cells(4, 17) Then
            greatest_volume = ws.Range("L" & a).Value
            ws.Range("R4").Value = ws.Range("I" & a).Value
            ws.Range("Q4").Value = greatest_volume
        End If
        
    ws.Range("R2", "R3").NumberFormat = "0.00%" 'friend michael told me this
    
            
    Next a
    
Next ws
    
End Sub
