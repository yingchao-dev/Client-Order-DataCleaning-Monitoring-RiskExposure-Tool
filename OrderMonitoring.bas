Attribute VB_Name = "OrderMonitoring"
Option Explicit

Sub MonitorPendingOrders()

    Dim thresholdMinutes As Long
    thresholdMinutes = 30

    On Error Resume Next
    Worksheets("Pending_Monitoring").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Pending_Monitoring"
    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")
    
    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 6)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Order ID"
    wsMonitor.Cells(1, 2).Value = "ISIN"
    wsMonitor.Cells(1, 3).Value = "Client ID"
    wsMonitor.Cells(1, 4).Value = "Timestamp"
    wsMonitor.Cells(1, 5).Value = "Pending Minutes"
    wsMonitor.Cells(1, 6).Value = "Alert"

    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1

    Dim orderAgeMinutes As Double
    
    'Zu Testzwecken habe ich die Schleife hier auf den Bereich 2 bis 100 begrenzt
    For i = 2 To 100

        If wsCleaned.Cells(i, 8).Value = "Pending" Then

            If IsDate(wsCleaned.Cells(i, 3).Value) Then
                orderAgeMinutes = DateDiff("n", wsCleaned.Cells(i, 3).Value, Now())

                If orderAgeMinutes > thresholdMinutes Then
                    monitorRow = monitorRow + 1
                    wsMonitor.Cells(monitorRow, 1).Value = wsCleaned.Cells(i, 1).Value
                    wsMonitor.Cells(monitorRow, 2).Value = wsCleaned.Cells(i, 2).Value
                    wsMonitor.Cells(monitorRow, 3).Value = wsCleaned.Cells(i, 9).Value
                    wsMonitor.Cells(monitorRow, 4).Value = wsCleaned.Cells(i, 3).Value
                    wsMonitor.Cells(monitorRow, 5).Value = orderAgeMinutes
                    wsMonitor.Cells(monitorRow, 6).Value = "Pending > " & thresholdMinutes & " min"
                End If
            End If

        End If

    Next i

    wsMonitor.Columns.AutoFit

End Sub

Sub MonitorPriceDeviation()

    On Error Resume Next
    Worksheets("Price_Deviation_Monitoring").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Price_Deviation_Monitoring"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 6)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Order ID"
    wsMonitor.Cells(1, 2).Value = "ISIN"
    wsMonitor.Cells(1, 3).Value = "Execution Price"
    wsMonitor.Cells(1, 4).Value = "Baseline Price"
    wsMonitor.Cells(1, 5).Value = "Deviation %"
    wsMonitor.Cells(1, 6).Value = "Alert"

    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1

    Dim dictSum As Object
    Dim dictCount As Object
    Dim isin As Variant
    Dim ep As Variant
    Set dictSum = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")

    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        ep = wsCleaned.Cells(i, 11).Value

        If Not IsEmpty(ep) And wsCleaned.Cells(i, 12).Value = "Yes" Then
            If dictSum.Exists(isin) Then
                dictSum(isin) = dictSum(isin) + ep
                dictCount(isin) = dictCount(isin) + 1
            Else
                dictSum(isin) = ep
                dictCount(isin) = 1
            End If
        End If
    Next i

    Dim avgDict As Object
    Set avgDict = CreateObject("Scripting.Dictionary")

    Dim key As Variant
    For Each key In dictSum.Keys
        avgDict(key) = dictSum(key) / dictCount(key)
    Next key

    ' Fester Schwellenwert als vereinfachter Ansatz - siehe README, Abschnitt "2.1 Preisabweichungserkennung", fuer Erlaeuterung der Limitationen.
    Dim deviation As Double
    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        ep = wsCleaned.Cells(i, 11).Value

        If Not IsEmpty(ep) And wsCleaned.Cells(i, 12).Value = "Yes" Then
            deviation = (ep - avgDict(isin)) / avgDict(isin)

            If Abs(deviation) > 0.02 Then
                monitorRow = monitorRow + 1
                wsMonitor.Cells(monitorRow, 1).Value = wsCleaned.Cells(i, 1).Value
                wsMonitor.Cells(monitorRow, 2).Value = wsCleaned.Cells(i, 2).Value
                wsMonitor.Cells(monitorRow, 3).Value = wsCleaned.Cells(i, 11).Value
                wsMonitor.Cells(monitorRow, 4).Value = avgDict(isin)
                wsMonitor.Cells(monitorRow, 5).Value = deviation
                wsMonitor.Cells(monitorRow, 5).NumberFormat = "0.00%"
                wsMonitor.Cells(monitorRow, 6).Value = "Preisabweichung uber 2%"
            End If
        End If
    Next i

    wsMonitor.Columns.AutoFit

End Sub
