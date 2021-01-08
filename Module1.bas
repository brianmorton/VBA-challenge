Attribute VB_Name = "Module1"
'declare variables

Dim r As Long
Public nxtticker As Long
Dim ticcomp As String
Dim ticker As String
Dim Rtotal As Double
Dim Lastrow As Double

Dim cl As Double
Dim op As Double


Sub forloop()
Lastrow = Range("A" & Rows.Count).End(xlUp).Row
nxtticker = 2
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percentage Change"
Cells(1, 12) = "Total Stock Volume"
For r = 2 To Lastrow


'ticker rules

If StrComp(ticker, ticcomp) = 0 Or Null Then
ticker = Cells(r, 1).Value
ticcomp = ticker

'yearchange instructions
If op = 0 Then
op = Cells(r, 3)
End If
cl = Cells(r, 6)


'total instructions
Rtotal = Rtotal + Cells(r, 7).Value

Else

'assign variables to cell
Cells(nxtticker, 9).Value = ticcomp
Cells(nxtticker, 10).Value = cl - op
Cells(nxtticker, 12).Value = Rtotal
'If statement to catch zeros and assign

If Cells(nxtticker, 10) = 0 Then
Cells(nxtticker, 11) = 0
Else
Cells(nxtticker, 11).Value = ((cl - op) / op)
End If


'format colors
If Cells(nxtticker, 10) > 0 Then
Cells(nxtticker, 10).Interior.ColorIndex = 4

ElseIf Cells(nxtticker, 10) < 0 Then
Cells(nxtticker, 10).Interior.ColorIndex = 3

'format percentage
Cells(nxtticker, 11).NumberFormat = "0.00%"

End If

'clean variables for next loop
ticcomp = ticker
op = 0
cl = 0
Rtotal = 0
nxtticker = nxtticker + 1
 
End If

'Set ticker to next value
ticker = Cells(r + 1, 1).Value
Next r

End Sub

