Attribute VB_Name = "Module7"
Sub stockAnalysis_bonus()
'Find and output the Greatest % Increase, % Decrease, and Total Volume

' Define variables
Dim r As Integer
Dim blankRow As Double
Dim greatestInc As String
Dim greatestIncNum As Double
Dim greatestDec As String
Dim greatestDecNum As Double
Dim greatestVol As String
Dim greatestVolNum As Double

' Discover final row
blankRow = Cells(Rows.Count, 9).End(xlUp).Row

' Name output cells
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

    ' Find maximum percent change
        greatestIncNum = WorksheetFunction.Max(Range("K2:K" & blankRow))
    
        ' Output Greatest % Increase Value
            Range("Q2").Value = greatestIncNum
    
    ' Find minimum percent change
        greatestDecNum = WorksheetFunction.Min(Range("K2:K" & blankRow))
    
        ' Output Greatest % Decrease Value
            Range("Q3").Value = greatestDecNum
    
    ' Find maximum volume
        greatestVolNum = WorksheetFunction.Max(Range("L2:L" & blankRow))

        ' Output Greatest Total Volume Value
            Range("Q4").Value = greatestVolNum
    
' Change number format of cells
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0,000"
    
' Loop through data to find ticker that corresponds with the summarized values
For r = 2 To blankRow

    ' If % change is Greatest % Increase Value, output the ticker into the appropriate cell
    If Cells(r, 11).Value = greatestIncNum Then
        Range("P2").Value = Cells(r, 9).Value
    
      ' If % change is Greatest % Decrease Value, output the ticker into the appropriate cell
    ElseIf Cells(r, 11).Value = greatestDecNum Then
        Range("P3").Value = Cells(r, 9).Value
    End If
    
      ' If volume is Greatest Total Volume Value, output the ticker into the appropriate cell
    If Cells(r, 12).Value = greatestVolNum Then
        Range("P4").Value = Cells(r, 9).Value
    End If
    
' Move to next row
Next r

' Autofit columns
Columns("O:Q").AutoFit

End Sub
