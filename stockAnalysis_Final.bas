Attribute VB_Name = "Module2"
Sub stockAnalysisFinal():
'Best & working code (3/23/21)
'Consolidated ticker & volume with yearly change & percent change in one if statement

' Define variables
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim volume As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim blankrow As Double
    Dim outputBlankRow As Double
 
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
    n = 2
    
' Loop through data to designate tickers, volume, year change & percent change
For r = 2 To blankrow

    ' If the next ticker is different (changed tickers)
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        
        ' TICKER:
        ' Store ticker name
            ticker = Cells(r, 1).Value
        
        ' Output ticker name into sheet
            Cells(n, 9).Value = ticker
        
        'VOLUME:
        ' Add up remaining volume
            volume = volume + Cells(r, 7)
        
        ' Output total volume
            Cells(n, 12).Value = volume
        
        ' Reset volume
            volume = 0
        
        'CLOSE PRICE:
        ' Store Closing Price for current ticker
            close_price = Cells(r, 6).Value

        'YEARLY CHANGE:
        ' Calculate yearly change
            year_change = Round(close_price - open_price, 2)
        
        ' Output Yearly Change
            Cells(n, 10).Value = year_change
            
            'PERCENT CHANGE:
            ' If open price doesn't equal 0
                If open_price <> 0 Then
            
                    ' Calculate percent change
                    percent_change = year_change / open_price
            
            ' If open price equals 0
                ElseIf open_price = 0 Then
            
                    ' Set percent change to 0
                    percent_change = 0
                
                End If

        ' Output Percent Change
            Cells(n, 11).Value = percent_change
        
        ' Reset Yearly Change
            year_change = 0
        
            ' TEST : Output Closing Price
            'Cells(n, 14).Value = close_price
            
            ' TEST: Output Opening Price
            'Cells(n, 15).Value = open_price
            
        ' Update step counter for output
            n = n + 1
    
    ' If the previous ticker is different (changed tickers)
        ElseIf Cells(r, 1).Value <> Cells(r - 1, 1).Value Then
        
        'OPEN PRICE:
        ' Store Open Price for current ticker
            open_price = Cells(r, 3).Value
            
        'VOLUME:
        ' Add volume of cell to volume
            volume = volume + Cells(r, 7)
            
    ' If the next ticker name is the same
        ElseIf Cells(r, 1).Value = Cells(r + 1, 1).Value Then

        'VOLUME:
        ' Add volume of cell to volume
            volume = volume + Cells(r, 7)

        End If
    
' Move to next row
Next r

'CONDITIONAL FORMATTING:
' Define new blank row for output columns
outputBlankRow = Cells(Rows.Count, 10).End(xlUp).Row

' Loop through output data to set color formatting
For r = 2 To outputBlankRow

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

' Add percent number formatting for "percent change"
Columns("K").NumberFormat = "0.00%"

'Add number formatting for volume
Columns("L").NumberFormat = "0,000"

End Sub

