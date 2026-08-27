Attribute VB_Name = "OrderDataCleaning"
Option Explicit

Function ParsePrice(txt As Variant) As Variant
    If IsEmpty(txt) Or Trim(CStr(txt)) = "" Then
       ParsePrice = Null
        Exit Function
    End If
    Dim cleanTxt As String
    cleanTxt = Trim(CStr(txt))
    If cleanTxt = "N/A" Or cleanTxt = "NULL" Or cleanTxt = "-" Or cleanTxt = "#REF!" Or cleanTxt = "???" Or cleanTxt = "###" Then
       ParsePrice = Null
        Exit Function
    End If
    If InStr(cleanTxt, "�") > 0 Then
        cleanTxt = Replace(cleanTxt, "�", "")
    End If
    Dim lastComma As Long
    Dim lastDot As Long
    lastComma = InStrRev(cleanTxt, ",")
    lastDot = InStrRev(cleanTxt, ".")

    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            cleanTxt = Replace(cleanTxt, ".", "")
            cleanTxt = Replace(cleanTxt, ",", ".")
        Else
            cleanTxt = Replace(cleanTxt, ",", "")
        End If
    ElseIf lastComma > 0 Then
        cleanTxt = Replace(cleanTxt, ",", ".")
    End If
    If cleanTxt <> "" Then
        ParsePrice = Val(cleanTxt)
    Else
        ParsePrice = Null
    End If
End Function

Function ParseQuantity(txt As Variant) As Variant
If IsEmpty(txt) Or Trim(txt) = "" Or Not IsNumeric(txt) Then
   ParseQuantity = Null
   Exit Function
End If
ParseQuantity = CLng(txt)
End Function

Function ParseTimestamp(txt As Variant) As Variant
    If IsDate(txt) Then
        ParseTimestamp = CDate(txt)
    Else
        ParseTimestamp = Null
    End If
End Function

Function ValidateISINFormat(txt As Variant) As Boolean
Dim cleaned As String
If IsEmpty(txt) Or Trim(txt) = "" Then
   ValidateISINFormat = False
   Exit Function
End If
cleaned = Trim(txt)
If Len(cleaned) <> 12 Then
   ValidateISINFormat = False
   Exit Function
End If
If Not (cleaned Like "[A-Z][A-Z][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]") Then
    ValidateISINFormat = False
    Exit Function
End If
ValidateISINFormat = True
End Function

Function ValidateISINInSecurityMaster(isin As String, secMasterRange As Range) As Boolean
Dim found As Range
Set found = secMasterRange.Find(what:=isin, LookAt:=xlWhole)
If found Is Nothing Then
   ValidateISINInSecurityMaster = False
Else
   ValidateISINInSecurityMaster = True
End If
End Function

Function ValidateOrderIDFormat(txt As Variant) As Boolean
Dim cleaned As String
If IsEmpty(txt) Or Trim(txt) = "" Then
   ValidateOrderIDFormat = False
   Exit Function
End If
cleaned = Trim(txt)
If Not (cleaned Like "ORD########") Then
    ValidateOrderIDFormat = False
    Exit Function
End If
ValidateOrderIDFormat = True
End Function

Function ValidateClientIDFormat(txt As Variant) As Boolean
Dim cleaned As String
If IsEmpty(txt) Or Trim(txt) = "" Then
   ValidateClientIDFormat = False
   Exit Function
End If
cleaned = Trim(txt)
If Not (cleaned Like "[A-Z][A-Z]####") Then
    ValidateClientIDFormat = False
    Exit Function
End If
ValidateClientIDFormat = True
End Function

Function ValidateClientIDClientMaster(ClientID As String, clientMasterRange As Range) As Boolean
Dim found As Range
Set found = clientMasterRange.Find(what:=ClientID, LookAt:=xlWhole)
If found Is Nothing Then
   ValidateClientIDClientMaster = False
Else
   ValidateClientIDClientMaster = True
End If
End Function

Function ValidateStatusFormat(txt As Variant) As Boolean
Dim cleaned As String
If IsEmpty(txt) Or Trim(txt) = "" Then
   ValidateStatusFormat = False
   Exit Function
End If
cleaned = UCase(Trim(txt))
If cleaned = "PENDING" Or cleaned = "PARTIALLY EXECUTED" Or cleaned = "EXECUTED" Or cleaned = "CANCELLED" Or cleaned = "REJECTED" Then
   ValidateStatusFormat = True
Else
   ValidateStatusFormat = False
End If
End Function

Function ValidateMarketPhaseFormat(txt As Variant) As Boolean
Dim cleaned As String
If IsEmpty(txt) Or Trim(txt) = "" Then
   ValidateMarketPhaseFormat = False
   Exit Function
End If
cleaned = UCase(Trim(txt))
If cleaned = "CONTINUOUS" Or cleaned = "AUCTION" Or cleaned = "CIRCUIT BREAKER" Then
   ValidateMarketPhaseFormat = True
Else
   ValidateMarketPhaseFormat = False
End If
End Function

Sub Test()

On Error Resume Next
Worksheets("Data_Quality").Delete
Worksheets("Cleaned_Data").Delete
On Error GoTo 0

Dim wsQuality As Worksheet
Set wsQuality = Worksheets.Add
wsQuality.Name = "Data_Quality"

Dim wsCleaned As Worksheet
Set wsCleaned = Worksheets.Add
wsCleaned.Name = "Cleaned_Data"
Sheets("Raw_Client_Orders").Rows(1).Copy Destination:=wsCleaned.Rows(1)
wsCleaned.Cells(1, 11).Copy
wsCleaned.Cells(1, 12).PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
wsCleaned.Cells(1, 12).Value = "Eligible_For_Price_Check"

wsQuality.Cells(1, 1).Value = "Row"
wsQuality.Cells(1, 2).Value = "OrderID"
wsQuality.Cells(1, 3).Value = "Problem"

' 工作表引用
Dim wsData As Worksheet
Dim wsSMData As Worksheet
Dim wsCMData As Worksheet

' 参考表查找范围
Dim secMasterRange As Range
Dim cliMasterRange As Range

Set wsData = Worksheets("Raw_Client_Orders")
Set wsSMData = Worksheets("Security_Master")
Set wsCMData = Worksheets("Client_Master")
Set secMasterRange = wsSMData.Range("A:A")
Set cliMasterRange = wsCMData.Range("A:A")

' 循环控制与计数
Dim i As Long
Dim lastRow As Long
Dim qualityRow As Long
Dim rowNum As Long

' 字段清洗/解析结果
Dim parsed As Variant
Dim parsedLimitPrice As Variant
Dim parsedExecutionPrice As Variant
Dim parsedQty As Variant

' 字段格式/存在性校验结果
Dim isISINValid As Boolean
Dim isISINFound As Boolean
Dim isClientIDValid As Boolean
Dim isClientIDFound As Boolean
Dim isOrderIDValid As Boolean
Dim isMarketPhaseValid As Boolean

' 原始值与清洗后的字符串
Dim orderID As String
Dim rawISIN As String
Dim cleanedISIN As String
Dim rawStatus As String
Dim rawClientID As String
Dim cleanedClientID As String
Dim orderType As String

' 其他工具变量
Dim dict As Object
Dim key As Variant
Dim rowList As String

Set dict = CreateObject("Scripting.Dictionary")

qualityRow = 1
lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

orderID = wsData.Cells(i, 1).Value
parsed = ParseTimestamp(wsData.Cells(i, 3).Value)
parsedLimitPrice = ParsePrice(wsData.Cells(i, 7).Value)
parsedExecutionPrice = ParsePrice(wsData.Cells(i, 11).Value)
parsedQty = ParseQuantity(wsData.Cells(i, 4).Value)
isMarketPhaseValid = ValidateMarketPhaseFormat(wsData.Cells(i, 10).Value)
isOrderIDValid = ValidateOrderIDFormat(wsData.Cells(i, 1).Value)

rawISIN = wsData.Cells(i, 2).Value
cleanedISIN = Trim(rawISIN)
rawStatus = wsData.Cells(i, 8).Value
rawClientID = wsData.Cells(i, 9).Value
cleanedClientID = Trim(rawClientID)

isISINValid = ValidateISINFormat(rawISIN)
If Not isISINValid Then
   qualityRow = qualityRow + 1
    wsQuality.Cells(qualityRow, 1).Value = i
    wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
    wsQuality.Cells(qualityRow, 3).Value = "ISIN is empty or invalid ISIN format"
Else
    isISINFound = ValidateISINInSecurityMaster(cleanedISIN, secMasterRange)
    If Not isISINFound Then
       qualityRow = qualityRow + 1
       wsQuality.Cells(qualityRow, 1).Value = i
       wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
       wsQuality.Cells(qualityRow, 3).Value = "ISIN not found in Security_Master"
    Else
       wsCleaned.Cells(i, 2).Value = cleanedISIN
    End If
End If

isClientIDValid = ValidateClientIDFormat(rawClientID)
If Not isClientIDValid Then
   qualityRow = qualityRow + 1
    wsQuality.Cells(qualityRow, 1).Value = i
    wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
    wsQuality.Cells(qualityRow, 3).Value = "ClientID is empty or invalid ClientID format"
Else
    isClientIDFound = ValidateClientIDClientMaster(cleanedClientID, cliMasterRange)
    If Not isClientIDFound Then
       qualityRow = qualityRow + 1
       wsQuality.Cells(qualityRow, 1).Value = i
       wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
       wsQuality.Cells(qualityRow, 3).Value = "ClientID not found in Client_Master"
    Else
       wsCleaned.Cells(i, 9).Value = cleanedClientID
    End If
End If

If Not isMarketPhaseValid Then
   qualityRow = qualityRow + 1
   wsQuality.Cells(qualityRow, 1).Value = i
   wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
   wsQuality.Cells(qualityRow, 3).Value = "MarketPhase is empty or invalid"
Else
   wsCleaned.Cells(i, 10).Value = wsData.Cells(i, 10).Value
   If UCase(Trim(wsData.Cells(i, 10).Value)) = "CONTINUOUS" Then
      wsCleaned.Cells(i, 12).Value = "Yes"
    Else
        wsCleaned.Cells(i, 12).Value = "No"
    End If
End If

If Not ValidateStatusFormat(rawStatus) Then
   qualityRow = qualityRow + 1
   wsQuality.Cells(qualityRow, 1).Value = i
   wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
   wsQuality.Cells(qualityRow, 3).Value = "Status is empty or invalid"
Else
   wsCleaned.Cells(i, 8).Value = wsData.Cells(i, 8).Value
End If

  If IsNull(parsedQty) Then
       qualityRow = qualityRow + 1
       wsQuality.Cells(qualityRow, 1).Value = i
       wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
       wsQuality.Cells(qualityRow, 3).Value = "Quantity is empty or not a number"
  ElseIf parsedQty < 0 Then
      wsData.Rows(i).Interior.Color = vbRed
      qualityRow = qualityRow + 1
      wsQuality.Cells(qualityRow, 1).Value = i
      wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
      wsQuality.Cells(qualityRow, 3).Value = "Negative Quantity"
  Else
      wsCleaned.Cells(i, 4).Value = parsedQty
  End If
  
  If IsNull(parsed) Then
     qualityRow = qualityRow + 1
     If IsEmpty(wsData.Cells(i, 3).Value) Or wsData.Cells(i, 3).Value = "" Or Trim(CStr(wsData.Cells(i, 3).Value)) = "" Then
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Timestamp is empty"
     Else
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Achtung: Unbekanntes Datenformat"
     End If
   Else
       wsCleaned.Cells(i, 3).Value = parsed
       wsCleaned.Cells(i, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
   End If

orderType = Trim(CStr(wsData.Cells(i, 6).Value))

If UCase(orderType) = "LIMIT ORDER" Then
    If IsNull(parsedLimitPrice) Then
        qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "LimitPrice is empty or invalid"
    ElseIf parsedLimitPrice = 0 Then
        qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "In der Spalte Limit Price koennten Parsing-Fehler auftreten"
    Else
        wsCleaned.Cells(i, 7).Value = parsedLimitPrice
    End If
ElseIf UCase(orderType) = "MARKET ORDER" Then
    If Not IsNull(parsedLimitPrice) Then
        qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Bei Market-Orders sollte kein Limitpreis vorhanden sein"
    End If
End If
  
  If UCase(Trim(rawStatus)) = "EXECUTED" Or UCase(Trim(rawStatus)) = "PARTIALLY EXECUTED" Then
     If IsNull(parsedExecutionPrice) Then
        qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Executed but Execution Price is empty"
      Else
          wsCleaned.Cells(i, 11).Value = parsedExecutionPrice
     End If
  Else
      If Not IsNull(parsedExecutionPrice) Then
         qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "Execution Price sollte nicht vorhanden sein"
      End If
  End If
  
  If Not isOrderIDValid Then
        qualityRow = qualityRow + 1
        wsQuality.Cells(qualityRow, 1).Value = i
        wsQuality.Cells(qualityRow, 2).Value = wsData.Cells(i, 1).Value
        wsQuality.Cells(qualityRow, 3).Value = "OrderID is empty or invalid"
   Else
     If dict.Exists(orderID) Then
        dict(orderID) = dict(orderID) & "," & i
     Else
        dict(orderID) = i
     End If
  End If
  
  wsCleaned.Cells(i, 5).Value = wsData.Cells(i, 5).Value
  wsCleaned.Cells(i, 6).Value = wsData.Cells(i, 6).Value
Next i

For Each key In dict.Keys
 If InStr(dict(key), ",") > 0 Then
    qualityRow = qualityRow + 1
    rowList = Replace(dict(key), ",", ", ")
    wsQuality.Cells(qualityRow, 1).NumberFormat = "@"
    wsQuality.Cells(qualityRow, 1).Value = dict(key)
    wsQuality.Cells(qualityRow, 2).Value = key
    wsQuality.Cells(qualityRow, 3).Value = "doppelte Bestellnummer - betroffene Zeilen: " & rowList
 Else
    rowNum = CLng(dict(key))
    wsCleaned.Cells(rowNum, 1).Value = wsData.Cells(rowNum, 1).Value
 End If
Next key

wsCleaned.Columns.AutoFit
wsQuality.Columns.AutoFit

End Sub
