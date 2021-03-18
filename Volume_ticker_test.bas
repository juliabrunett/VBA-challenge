Attribute VB_Name = "Module2"
Sub volume_ticker_test():

' Define variables
    Dim volume As Double
    Dim r As Double
    Dim i As Long
    Dim n As Long
    Dim ticker As String
    
' Initialize variables
    volume = 0
    n = 2
    i = 2

' Loop through data to sum volume & provide ticker
For r = 2 To 70926

' If the ticker names are different
    If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
    
        ' Store ticker name
        ticker = Cells(r, 1).Value
        
        ' Output ticker name into sheet
        Cells(i, 9).Value = ticker
        
        ' Add up remaining volume
        volume = volume + Cells(r, 7)
        
        ' Output total volume
        Cells(i, 13).Value = volume
        
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

End Sub

