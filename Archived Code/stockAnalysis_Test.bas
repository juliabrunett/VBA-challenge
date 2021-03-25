Attribute VB_Name = "Module6"
Sub stockAnalysis2():
'3/21/21 - most recent success
'Working code (unconsolidated)
'Not intended for use in final analysis

' Define variables
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim volume As Double
    Dim year_change As Double
    Dim blankrow As Double
    Dim percent_change As Double
 
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
    
' Loop through data to designate tickers, volume, year change & percent change
For r = 2 To blankrow

    ' If the next ticker is different (changed tickers)
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
        ' Store Closing Price for current ticker
            close_price = Cells(r, 6).Value

    ' Calculate yearly change & percent change
        year_change = Round(close_price - open_price, 2)
        
    ' Output Year Change
        Cells(n, 10).Value = year_change
            
            ' If open price doesn't equal 0
            If open_price <> 0 Then
            
                ' Calculate percent change
                percent_change = (year_change / open_price) * 100
            
            ' If open price equals 0
            ElseIf open_price = 0 Then
            
                ' Set percent change to 0
                percent_change = 0
                
            End If

    ' Output Percent Change
        Cells(n, 11).Value = percent_change & "%"
        
    ' Reset Year Change
    year_change = 0
        
        ' TEST : Output Closing Price
            'Cells(n, 14).Value = close_price
            
        ' TEST: Output Opening Price
            'Cells(n, 15).Value = open_price
            
    ' Update step counter for output
        n = n + 1
    
' If date is greater than the minimum date and it is a different ticker from the previous
    ElseIf Cells(r, 1).Value <> Cells(r - 1, 1).Value Then
        
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

' Autofit columns
Columns("I:L").AutoFit

End Sub

