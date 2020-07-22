Sub Stock_Market()

        'Set variable for Ticker
        Dim Ticker_Symbol As String
        
        'Set Variable for Yearly Change
        
        Dim Yearly_Change As Double
        
        Yearly_Change = 0
        
        'Set Variable for StockV
        Dim Stock_Vol As Double
        
        Stock_Vol = 0
        
        'Set Ticker Integer
        Dim Tick_Name As Integer
        
        'Set Variable for Last Row
        
        Dim Last_Row As Long
        
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set Variable for Open Price
        
        Dim Open_Price As Double
        
        Open_Price = 0
        
        Open_Price = Cells(2, 3).Value
        
        ' Set Varible for Closing
        
        Dim Close_Price As Double
        
        Close_Price = 0
        
        'Set Percent Change
        
        Dim Percent_Change As Double
        
        Percent_Change = 0
        
        'Format Percentage
        
        Range("K2:K" & Last_Row).NumberFormat = "0.00%"
        
        'Add info to table
        
        Info_Print = 2
        
        'Headers
        
          Range("I1").Value = "Ticker"
          Range("J1").Value = "Yearly Change"
          Range("K1").Value = "Percent Change"
          Range("L1").Value = "Total Stock Volume"
        
        'Loop through all Stocks
        
    For i = 2 To Last_Row
    
    'Populate Ticker Symbol
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set Ticker Name
    
            Ticker_Symbol = Cells(i, 1).Value
    
        'Print Yearly percent change
    
            Close_Price = Cells(i, 6).Value
    
            Yearly_Change = Close_Price - Open_Price
    
                If Open_Price <> 0 Then
    
                    Percent_Change = (Yearly_Change / Open_Price)
    
                'Print Range J and K
    
                        Range("J" & Info_Print).Value = Yearly_Change
    
                        Range("K" & Info_Print).Value = Percent_Change
                     
                     'Add color on yearly percent change
        
                    If Yearly_Change > 0 Then
                    'Make # Green
                    Range("J" & Info_Print).Interior.Color = RGB(0, 128, 0)
                     Else
                    Range("J" & Info_Print).Interior.Color = RGB(255, 0, 0)
                    End If
                
                End If
                
                
                'Reset Change for next row value
    
                    Yearly_Change = 0
    
                    Close_Price = 0
    
                    Open_Price = Cells(i + 1, 3).Value
    
    
                'Stock Volume Total
    
            Stock_Vol = Stock_Vol + Cells(i, 7).Value
    
                 'Print TickerSymbol to Column I
    
                    Range("I" & Info_Print).Value = Ticker_Symbol
    
                    'Print Stock Volume Total
    
                    Range("L" & Info_Print).Value = Stock_Vol
    
        'Add to next row
    
            Info_Print = Info_Print + 1
    
        'Reset for next Ticker
    
             Stock_Vol = 0
    
        'If names are same add to Stock Vol
    
        Else
        
            Stock_Vol = Stock_Vol + Cells(i, 7).Value
    
    
        End If
    
    Next i
End Sub
