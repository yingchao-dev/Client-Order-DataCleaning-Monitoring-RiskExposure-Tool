Attribute VB_Name = "OrderDataCleaning"
Option Explicit

Sub Test()

On Error Resume Next
Worksheets("Data_Quality").Delete
On Error GoTo 0

Dim wsQuality As Worksheet
Set wsQuality = Worksheets.Add
wsQuality.Name = "Data_Quality"

wsQuality.Cells(1, 1).Value = "Row"
wsQuality.Cells(1, 2).Value = "OrderID"
wsQuality.Cells(1, 3).Value = "Problem"

Dim qualityRow As Long
Dim i As Long
Dim lastRow As Long
Dim wsData As Worksheet
Set wsData = Worksheets("Sheet1")

qualityRow = 1
lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
  If wsData.Cells(i, 2).Value < 0 Then
     qualityRow = qualityRow + 1
     wsQuality.Cells(qualityRow, 1).Value = i
     wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
     wsQuality.Cells(qualityRow, 3).Value = "Negative Quantity"
  End If
  If IsEmpty(wsData.Cells(i, 2).Value) Then
  qualityRow = qualityRow + 1
  wsQuality.Cells(qualityRow, 1).Value = i
  wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
  wsQuality.Cells(qualityRow, 3).Value = "Quantity is empty"
  End If
  If IsEmpty(wsData.Cells(i, 3).Value) Then
  qualityRow = qualityRow + 1
  wsQuality.Cells(qualityRow, 1).Value = i
  wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
  wsQuality.Cells(qualityRow, 3).Value = "Price is empty"
  End If
  If IsEmpty(wsData.Cells(i, 5).Value) And wsData.Cells(i, 4) = "Executed" Then
  qualityRow = qualityRow + 1
  wsQuality.Cells(qualityRow, 1).Value = i
  wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
  wsQuality.Cells(qualityRow, 3).Value = "Executed but Execution Price is empty"
  End If
  
Next i

End Sub
