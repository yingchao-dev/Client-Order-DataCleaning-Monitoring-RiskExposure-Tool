Attribute VB_Name = "Module1"
Option Explicit
Sub Test()

Dim wsQuality As Worksheet
Set wsQuality = Worksheets.Add
wsQuality.Name = "Data_Quality"

wsQuality.Cells(1, 1).Value = "Row"
wsQuality.Cells(1, 2).Value = "OrderID"
wsQuality.Cells(1, 3).Value = "Problem"

Dim row As Long
Dim i As Long
row = 1
For i = 2 To 7
   If Worksheets("Sheet1").Cells(i, 2).Value < 0 Then
      row = row + 1
      wsQuality.Cells(row, 1).Value = i
      wsQuality.Cells(row, 2).Value = Worksheets("Sheet1").Cells(i, 1).Value
      wsQuality.Cells(row, 3).Value = "Negative Quantity"
    End If
    If IsEmpty(Worksheets("Sheet1").Cells(i, 2).Value) Then
      row = row + 1
      wsQuality.Cells(row, 1).Value = i
      wsQuality.Cells(row, 2).Value = Worksheets("Sheet1").Cells(i, 1).Value
      wsQuality.Cells(row, 3).Value = "Quantity is empty"
    End If
    If IsEmpty(Worksheets("Sheet1").Cells(i, 3).Value) Then
      row = row + 1
      wsQuality.Cells(row, 1).Value = i
      wsQuality.Cells(row, 2).Value = Worksheets("Sheet1").Cells(i, 1).Value
      wsQuality.Cells(row, 3).Value = "Price is empty"
    End If
    If IsEmpty(Worksheets("Sheet1").Cells(i, 5).Value) And Worksheets("Sheet1").Cells(i, 4) = "Executed" Then
      row = row + 1
      wsQuality.Cells(row, 1).Value = i
      wsQuality.Cells(row, 2).Value = Worksheets("Sheet1").Cells(i, 1).Value
      wsQuality.Cells(row, 3).Value = "Executed but Execution Price is empty"
    End If
    
Next i

End Sub
