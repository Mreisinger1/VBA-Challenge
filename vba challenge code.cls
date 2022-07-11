VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Vbachallenge()
    Dim i As Long
    Dim ws As Worksheet
    Dim Ticker As String
    Dim AccIndex As Long
AccIndex = 1
'loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
'add column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
 'determine last row
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Lastrow
'populate ticker box
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    AccIndex = AccIndex + 1
    Tickers = ws.Cells(i, 1).Value
    ws.Cells(AccIndex, 9).Value = Tickers
'calculate yearly change
    Yearly_change = (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
    ws.Cells(AccIndex, 10).Value = Yearly_change
'calculate percent change
percent_change = (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / (ws.Cells(i, 6).Value)
ws.Cells(AccIndex, 11).Value = percent_change
ws.Cells(AccIndex, 11).Value = Format(percent_change, "percent")
End If
'calculate total stock volume
If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
Stock_volume = Stock_volume + ws.Cells(i + 1, 7).Value
ws.Cells(AccIndex + 1, 12).Value = Stock_volume
End If
    If ws.Cells(AccIndex, 10).Value < 0 Then
    ws.Cells(AccIndex, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(AccIndex, 10).Interior.ColorIndex = 4
End If
Next i
Next ws
End Sub


