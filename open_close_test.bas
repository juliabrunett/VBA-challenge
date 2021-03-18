Attribute VB_Name = "Module4"
 Sub open_close_test():

' Define variables
    Dim ticker As String
    Dim stock_date As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim volume As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim blankrow As Double

' Define row step counter
    Dim r As Long
' Define output step counter
    Dim i As Integer
' Define additional output step counter
    Dim n As Integer
    
    
' Discover final row
    blankrow = Cells(Rows.Count, 1).End(xlUp).Row

' Title the output cells
    Cells(1, 9).Value = "Stock Ticker"
    Cells(1, 10).Value = "Close"
    Cells(1, 11).Value = "Open"
    Cells(1, 12).Value = "Total Stock Volume"
    
' Initialize variables
    volume = 0
    i = 2
    n = 2
    
' Initialize opening price
    open_price = Cells(2, 3).Value

' Output initial opening price
    Cells(2, 11).Value = open_price
    
' Loop through data to sum volume & provide ticker
For r = 2 To blankrow

' If the ticker names are different
    If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
    
        ' Store ticker name
        ticker = Cells(r, 1).Value
        
        ' Output ticker name into sheet
        Cells(i, 9).Value = ticker
        
        ' Add up remaining volume
        volume = volume + Cells(r, 7)
        
        ' Output total volume
        Cells(i, 12).Value = volume
        
        ' Reset volume
        volume = 0
           
        ' Update step counter for output
        i = i + 1
    
     ' If specified date is greater than the next date in row (meaning it changed tickers)
        If Cells(r, 2).Value > Cells(r + 1, 1).Value & Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
            ' Store Closing Price
            close_price = Cells(r, 6).Value
            
            'Store Open Price
            open_price = Cells(r + 1, 3).Value
            
            ' TEST : Output Closing Price
            Cells(n, 10).Value = close_price
            
            ' TEST: Output Opening Price
            Cells(n + 1, 11).Value = open_price
            
            ' Update step counter for output
            n = n + 1
            
        ' Calculate yearly change & percent change
    'year_change = close_price - open_price
    'percent_change = (close_price - open_price) / open_price
  
        'Cells(i,10).Value = year_change
        'Cells(i,11).Value = percent_change
    
        End If
        
' If the ticker names are the same
    Else
    
        ' Add volume of cell to total volume
        volume = volume + Cells(r, 7)
  
        
    End If
    
' Move to next row
Next r
        

End Sub


