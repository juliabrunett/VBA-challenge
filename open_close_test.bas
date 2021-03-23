Attribute VB_Name = "Module4"
 Sub open_close_test():
 'Test code for finding & outputting close & open prices for each ticker
 'Not intended for use in final analysis

' Define variables
    Dim ticker As String
    Dim stock_date As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim blankrow As Double

' Define row step counter
    Dim r As Long
' Define output step counter
    Dim n As Integer
    
' Discover final row
    blankrow = Cells(Rows.Count, 1).End(xlUp).Row

' Title the output cells
    Cells(1, 9).Value = "Stock Ticker"
    Cells(1, 10).Value = "Close"
    Cells(1, 11).Value = "Open"
    
' Initialize variables
    n = 2

' Loop through data to sum volume & provide ticker
For r = 2 To blankrow

' If the ticker names are different
    If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
    
        ' Store ticker name
        ticker = Cells(r, 1).Value
        
        ' Output ticker name into sheet
        Cells(n, 9).Value = ticker
        
        ' Store Closing Price
        close_price = Cells(r, 6).Value
            
        ' Output Closing Price
        Cells(n, 10).Value = close_price
            
        ' Update step counter for output
        n = n + 1
        
' If the ticker names are the same
    ElseIf Cells(r, 1).Value <> Cells(r - 1, 1).Value Then
    
        ' Store Open Price
        open_price = Cells(r, 3).Value
           
        ' Output Opening Price
        Cells(n, 11).Value = open_price
            
        End If
    
' Move to next row
Next r

End Sub


