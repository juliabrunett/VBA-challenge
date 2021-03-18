Attribute VB_Name = "Module5"
 Sub year_change_test():

' Define variables
    Dim ticker As String
    Dim stock_date As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim open_date As Long
    Dim close_date As Long
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
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
' Initialize variables
    volume = 0
    i = 2
    n = 2
    open_date = "20160101"
    'Cells(2, 2).Value
    close_date = "20161230"
    'Cells(263, 2).Value
    
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
        If Cells(r, 2).Value = close_date & Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
            ' Store Closing Price for current ticker
            close_price = Cells(r, 6).Value
            
            ' Store Open Price for next ticker
            'open_price = Cells(r + 1, 3).Value
        
        ' If date is less than next date
         ElseIf Cells(r, 2).Value = open_date Then
                
                ' Store Open Price for current ticker
                open_price = Cells(r, 3).Value
          
        End If
          
         ' Calculate yearly change & percent change
            year_change = close_price - open_price
            percent_change = (close_price - open_price) / open_price
            
           ' Output Year Change
            Cells(n, 10).Value = year_change
       
            ' Output Percent Change
            Cells(n, 11).Value = percent_change
            
            ' TEST : Output Closing Price
            Cells(n, 14).Value = close_price
            
            ' TEST: Output Opening Price
            Cells(n, 15).Value = open_price
            
             ' Update step counter for output
            n = n + 1
            
' If the ticker names are the same
    Else
    
        ' Add volume of cell to total volume
        volume = volume + Cells(r, 7)
        
    End If
    
' Move to next row
Next r
        
End Sub
