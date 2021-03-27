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

For Each ws In Worksheets

' Discover final row
blankRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

' Name output cells
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

    ' Find maximum percent change
        greatestIncNum = WorksheetFunction.Max(ws.Range("K2:K" & blankRow))
    
        ' Output Greatest % Increase Value
            ws.Range("Q2").Value = greatestIncNum
    
    ' Find minimum percent change
        greatestDecNum = WorksheetFunction.Min(ws.Range("K2:K" & blankRow))
    
        ' Output Greatest % Decrease Value
            ws.Range("Q3").Value = greatestDecNum
    
    ' Find maximum volume
        greatestVolNum = WorksheetFunction.Max(ws.Range("L2:L" & blankRow))

        ' Output Greatest Total Volume Value
            ws.Range("Q4").Value = greatestVolNum
    
' Change number format of cells
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "0,000"
    
' Loop through data to find ticker that corresponds with the summarized values
For r = 2 To blankRow

    ' If % change is Greatest % Increase Value, output the ticker into the appropriate cell
    If ws.Cells(r, 11).Value = greatestIncNum Then
        ws.Range("P2").Value = ws.Cells(r, 9).Value
    
      ' If % change is Greatest % Decrease Value, output the ticker into the appropriate cell
    ElseIf ws.Cells(r, 11).Value = greatestDecNum Then
        ws.Range("P3").Value = ws.Cells(r, 9).Value
    End If
    
      ' If volume is Greatest Total Volume Value, output the ticker into the appropriate cell
    If ws.Cells(r, 12).Value = greatestVolNum Then
        ws.Range("P4").Value = ws.Cells(r, 9).Value
    End If
    
' Move to next row
Next r

' Autofit columns
ws.Columns("O:Q").AutoFit

Next ws
End Sub
