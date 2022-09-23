Attribute VB_Name = "Module1"
Sub VBA_Assignment():
' Assignement: Create a script that loops through all the stocks for one year and outputs the following information
' The Ticker symbol
' Yearly change from opening price to closing price
' The percent change from opening price to closing price
' The total stock volume

'loop through all worksheets
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate


'Declare variables
Dim i As Double
Dim j As Double
Dim tblSet As Integer
Dim newcell As Integer
Dim summarytbl As Integer
Dim lrow As Long
Dim open_price As Double
Dim close_price As Double
Dim yearly_Change As Double
Dim percent_change As Double
Dim tVolume

Dim ticker As String

'assigning initial info to varibales
summarytbl = 2
open_price = Cells(2, 3).Value
close_price = 0
tVolume = 0

 
' identifies the last non-blank cell in a row .End(xlup) starts at the last occupied
' cell in a row and goes.
lrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

'loop through the ticker colum)
For i = 2 To lrow

'set the stock volume
tVolume = tVolume + Cells(i, 7).Value

'check to see if the next cell equals the previous sell
If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then


'Capture the close price
close_price = Cells(i, 6).Value
yearly_Change = open_price - close_price
percent_change = (yearly_Change / open_price)

Cells(summarytbl, 10).Value = ticker
Cells(summarytbl, 11).Value = yearly_Change
Cells(summarytbl, 12).Value = percent_change
Cells(summarytbl, 13).Value = tVolume

'move to the next row in the summary table
summarytbl = summarytbl + 1
'reset tvolume (total stock volume) to zero to capture the next count
tVolume = 0
'reset open price to capture next opening price
open_price = Cells(i + 1, 3).Value

'if the previous line matches the next line set the ticker to that value
Else

ticker = Cells(i, 1).Value

End If
Next i

For i = 2 To lrow

 For j = 11 To 12
 
    If Cells(i, j).Value >= 0 And j = 11 Then
    Cells(i, j).Interior.ColorIndex = 4
    
    ElseIf Cells(i, j).Value < 0 And j = 11 Then
    Cells(i, j).Interior.ColorIndex = 3
    
    ElseIf j = 12 Then
    Cells(i, j).NumberFormat = "#.##%"
 
End If
Next j
Next i

ActiveSheet.UsedRange.EntireColumn.AutoFit

Next ws
 
End Sub
