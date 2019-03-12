Sub Stock()
'Variable Declarations
On Error Resume Next
Dim ticker As String
Dim volume As Double
Dim i As Double, j As Double
Dim ws As Worksheet
Dim LR As Double
Dim openP As Double
Dim closeP As Double
Dim PERCH As Double


'For Every Sheet
For Each ws In Worksheets

'Last Row of Column A
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Last Row of Column L
LL = ws.Cells(Rows.Count, 12).End(xlUp).Row
'Last Row of Column J
LJ = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Updating Row Header
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Setting initial value
volume = 0
openP = ws.Cells(2, 3).Value
closeP = 0
j = 2

'For loop for the rows

For i = 2 To LR

'If loop to check against next row and update final columns

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
closeP = ws.Cells(i, 6).Value

ws.Cells(j, 10).Value = volume
ws.Cells(j, 11).Value = closeP - openP
PERCH = ws.Cells(j, 11).Value / openP
ws.Cells(j, 12).Value = PERCH

'Color Formatting Column K

If ws.Cells(j, 11).Value >= 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 4
ElseIf ws.Cells(j, 11).Value < 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 3

End If

'Formatting the column F into Percentage Numbers

ws.Range("L2:L" & LL).NumberFormat = "0.00%"
openP = ws.Cells(i + 1, 3).Value

j = j + 1
volume = 0

Else

volume = volume + ws.Cells(i, 7).Value
closeP = ws.Cells(i, 6).Value

'Ending main loops

End If

Next i
'End of Moderate assignment

'Begining of Hard Assignment:
'Final values: Percent Increase, Percentage Decrease and Max Volume

Dim rngL As Range
Dim rngJ As Range
Dim rngLMax As Double
Dim rngLMin As Double
Dim rngJMax As Double
Dim k As Double

'Set range from which to determine largest value
Set rngL = ws.Range("L2:L" & LL)
Set rngJ = ws.Range("J2:J" & LJ)


'Worksheet function MAX returns the largest value in a range

rngLMax = ws.Application.WorksheetFunction.Max(rngL)
rngLMin = ws.Application.WorksheetFunction.Min(rngL)
rngJMax = ws.Application.WorksheetFunction.Max(rngJ)

'Displays largest value
ws.Cells(2, 16).Value = rngLMax
ws.Cells(3, 16).Value = rngLMin

ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = rngJMax

'Looping for final values

For k = 2 To LL
If ws.Cells(k, 12).Value = ws.Cells(2, 16).Value Then
            ws.Cells(2, 15).Value = ws.Cells(k, 9).Value

ElseIf ws.Cells(k, 12).Value = ws.Cells(3, 16).Value Then

            ws.Cells(3, 15).Value = ws.Cells(k, 9).Value

ElseIf ws.Cells(k, 10).Value = ws.Cells(4, 16).Value Then
            ws.Cells(4, 15).Value = ws.Cells(k, 9).Value


End If

Next k

'Making the same changes in all worksheets

Next ws

End Sub
