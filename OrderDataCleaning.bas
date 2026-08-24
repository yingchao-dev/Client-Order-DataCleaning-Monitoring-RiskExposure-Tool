Attribute VB_Name = "OrderDataCleaning"
Option Explicit
Function ParsePrice(txt As Variant) As Variant
    If IsEmpty(txt) Or Trim(txt) = "" Then
       ParsePrice = Null
       Exit Function
    End If
    If InStr(txt, "€") > 0 Then
       txt = Replace(txt, "€", "")
    End If
    If InStr(txt, ".") > 0 Then
       txt = Replace(txt, ".", "")
    End If
    If InStr(txt, ",") > 0 Then
       txt = Replace(txt, ",", ".")
    End If
    ParsePrice = CDbl(txt)
End Function

Function ParseTimestamp(txt As Variant) As Variant
    If IsDate(txt) Then
        ParseTimestamp = CDate(txt)
    Else
        ParseTimestamp = Null
    End If
End Function

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
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
Dim orderID As String
Dim parsed As Variant
Dim parsedPrice As Variant

qualityRow = 1
lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastRow
parsed = ParseTimestamp(wsData.Cells(i, 6).Value)
parsedPrice = ParsePrice(wsData.Cells(i, 3).Value)

  Err.Clear
  On Error Resume Next
  
  orderID = wsData.Cells(i, 1).Value
  If dict.Exists(orderID) Then
     dict(orderID) = dict(orderID) & "," & i
  Else
     dict(orderID) = i
  End If

  If IsNumeric(wsData.Cells(i, 2).Value) Then
    If wsData.Cells(i, 2).Value < 0 Then
       qualityRow = qualityRow + 1
       wsQuality.Cells(qualityRow, 1).Value = i
       wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
       wsQuality.Cells(qualityRow, 3).Value = "Negative Quantity"
     End If
  Else
    qualityRow = qualityRow + 1
     wsQuality.Cells(qualityRow, 1).Value = i
     wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
     wsQuality.Cells(qualityRow, 3).Value = "Quantity is not a number"
  End If
  If IsEmpty(wsData.Cells(i, 2).Value) Then
     qualityRow = qualityRow + 1
     wsQuality.Cells(qualityRow, 1).Value = i
     wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
     wsQuality.Cells(qualityRow, 3).Value = "Quantity is empty"
  End If
  If IsEmpty(wsData.Cells(i, 5).Value) And wsData.Cells(i, 4) = "Executed" Then
     qualityRow = qualityRow + 1
     wsQuality.Cells(qualityRow, 1).Value = i
     wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
     wsQuality.Cells(qualityRow, 3).Value = "Executed but Execution Price is empty"
  End If
  If IsNull(parsed) Then
     qualityRow = qualityRow + 1
     If IsEmpty(wsData.Cells(i, 6).Value) Or wsData.Cells(i, 6).Value = "" Or Trim(CStr(wsData.Cells(i, 6).Value)) = "" Then
     
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Timestamp is empty"
     Else
        wsData.Rows(i).Interior.Color = vbYellow
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Achtung: Unbekanntes Datenformat"
     End If
   End If
   If IsNull(parsedPrice) Then
      qualityRow = qualityRow + 1
     wsQuality.Cells(qualityRow, 1).Value = i
     wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
     wsQuality.Cells(qualityRow, 3).Value = "Price is empty"
   Else
     wsQuality.Cells(i, 5).Value = wsData.Cells(i, 3).Value
   End If
    
  If Err.Number <> 0 Then
     Debug.Print "Row " & i & ": " & Err.Number & " - " & Err.Description
  End If
  
On Error GoTo 0

Next i

Dim key As Variant
For Each key In dict.Keys
 If InStr(dict(key), ",") > 0 Then
    qualityRow = qualityRow + 1
    wsQuality.Cells(qualityRow, 1).Value = dict(key)
    wsQuality.Cells(qualityRow, 2).Value = key
    wsQuality.Cells(qualityRow, 3).Value = "doppelte Bestellnummer"
 End If
Next key

End Sub
