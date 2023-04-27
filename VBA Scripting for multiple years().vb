Sub MultipleYearStockData()

For Each ws In Worksheets

Dim WorksheetName As String
'Current row
Dim i As Long
'start row of ticker block
Dim j As Long
'index counter to fill Ticker row
Dim TickCount As Long
'Last row colum A
Dim LastRowA As Long
'lastrow colum I
Dim LastRowI As Long
'Variable for percent change calculation
Dim PerChange As Double
'variable for greatest increase calculation
Dim GreatIncr As Double
'variable for greatest decrease calculation
Dim GreatDecr As Double
'variable for greatest total volume
Dim GreatVol As Double
'get the WorksheetName
WorksheetName = ws.Name

'Create column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume)"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


'Set Ticker Counter to first row
    TickCount = 2

'set start row to 2
    j = 2

'finding the last non-blank cell in column A
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox ("Last row in column A is " & LastRowA)

'Loop through all rows
For i = 2 To LastRowA

'Checking if ticker name changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Write ticker in column I (#9)
    ws.Cells(TickCount, 9).Value = ws.Cells(i, 6).Value = ws.Cells(i, 1).Value

'calculate and write Yearly Change in column J (#10)
    ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'Conditional Formating column
    If ws.Cells(TickCount, 10).Value < 0 Then
    

'Setting cells background color to red
ws.Cells(TickCount, 10).Interior.ColorIndex = 3

Else

'Setting cells background color to green
    ws.Cells(TickCount, 10).Interior.ColorIndex = 4

End If

'Calculate and write percent change in column K (#11)
If ws.Cells(j, 3).Value <> 0 Then
PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)

'Percent formating
    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")

Else

ws.Cells(TickCount, 11).Value = Format(0, "Percent")


End If

'Calculate and write total volume in column L (#12)
    ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

'Increase TickCount by 1
    TickCount = TickCount + 1

'Set new start row of the ticker block
    j = i + 1

End If


Next i

'Find last non-blank cell in column I
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox ("Last row in column I is " & LastRowI)

'Prepare for summary
GreatVol = ws.Cells(2, 12).Value
GreatIncr = ws.Cells(2, 11).Value
GreatDecr = ws.Cells(2, 11).Value

'Loop for summary
For i = 2 To LastRowI

'for greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
If ws.Cells(i, 12).Value > GreatVol Then
GreatVol = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

Else

GreatVol = GreatVol

End If

'for greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
If ws.Cells(i, 11).Value > GreatIncr Then
GreatIncr = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

Else

GreatIncr = GreatIncr

End If

'for greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
If ws.Cells(i, 11).Value < GreatDecr Then
GreatDecr = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value = ws.Cells(i, 9).Value

Else

GreatDecr = GreatDecr

End If

'Write summary results in ws.Cells
ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")


Next i

'Adjust column width automatically
Worksheets(WorksheetName).Columns("A:Z").AutoFit

Next ws




End Sub
