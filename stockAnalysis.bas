Attribute VB_Name = "Module5"
Sub stockAnalysis():

' Define variables
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim volume As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim blankrow As Double
    Dim max_date As Double
    Dim min_date As Double
 
' Define row step counter
    Dim r As Long
' Define output step counter
    Dim i As Integer
' Define additional output step counter
    Dim n As Integer
    
' Define final row in ticker column
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
    
' Define minimum date & maximum date for stock
    max_date = WorksheetFunction.Max(Range("B2:B" & Range("B" & Rows.Count).End(xlUp).Row))
    min_date = WorksheetFunction.Min(Range("B2:B" & Range("B" & Rows.Count).End(xlUp).Row))

' Loop through data to sum volume & provide ticker
For r = 2 To blankrow

' If specified date is equal to the maximum date and the next ticker is different
     If Cells(r, 2).Value = max_date And Cells(r, 1).Value <> Cells(r + 1, 1).Value Or Cells(r, 2).Value < max_date And Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
    ' Store Closing Price for current ticker
        close_price = Cells(r, 6).Value

    ' Calculate yearly change & percent change
        year_change = close_price - open_price
        percent_change = (year_change / open_price) * 100
        
    ' Output Year Change
        Cells(n, 10).Value = year_change
        
    ' Reset Year Change
    year_change = 0
       
    ' Output Percent Change
        Cells(n, 11).Value = percent_change & "%"
    
    ' Reset Percent Change
    percent_change = 0
        
        ' TEST : Output Closing Price
            'Cells(n, 14).Value = close_price
            
        ' TEST: Output Opening Price
            'Cells(n, 15).Value = open_price
            
    ' Update step counter for output
        n = n + 1
        
' If date is equal to the minimum date
    ElseIf Cells(r, 2).Value = min_date Then
                
    ' Store Open Price for current ticker
        open_price = Cells(r, 3).Value
    
' If date is greater than the minimum date and it is a different ticker from the previous
    ElseIf Cells(r, 2).Value > min_date And Cells(r, 1).Value <> Cells(r - 1, 1).Value Then
        
    ' Store Open Price for current ticker
        open_price = Cells(r, 3).Value

    End If

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
            
    ' If the ticker names are the same
        Else
    
        ' Add volume of cell to total volume
            volume = volume + Cells(r, 7)

        End If
    
' Move to next row
Next r
        
' Define new blank row for output columns
outputBlankrow = Cells(Rows.Count, 10).End(xlUp).Row

' Loop through output data to set color formatting
For r = 2 To outputBlankrow

    ' Setting conditional formatting
        If Cells(r, 10).Value > 0 Then
        
            ' Set values over 0 to green
            Cells(r, 10).Interior.ColorIndex = 4
            
        ElseIf Cells(r, 10).Value <= 0 Then
        
            'Set values under or equal to 0 to red
            Cells(r, 10).Interior.ColorIndex = 3
            
        End If
        
' Move to next row
Next r

End Sub
