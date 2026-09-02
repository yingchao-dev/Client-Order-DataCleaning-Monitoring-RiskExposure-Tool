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

        If UCase(wsCleaned.Cells(i, 8).Value) = "PENDING" Then

            If Not IsEmpty(wsCleaned.Cells(i, 3).Value) Then
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
    Dim threshold As Double
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

    ' Fester Schwellenwert als vereinfachter Ansatz - siehe README, Abschnitt "Preisabweichungserkennung", fuer Erlaeuterung der Limitationen.
    Dim deviation As Double
    threshold = 0.05
    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        ep = wsCleaned.Cells(i, 11).Value
        
        If Not IsEmpty(ep) And wsCleaned.Cells(i, 12).Value = "Yes" Then
            deviation = (ep - avgDict(isin)) / avgDict(isin)

            If Abs(deviation) > threshold Then
                monitorRow = monitorRow + 1
                wsMonitor.Cells(monitorRow, 1).Value = wsCleaned.Cells(i, 1).Value
                wsMonitor.Cells(monitorRow, 2).Value = wsCleaned.Cells(i, 2).Value
                wsMonitor.Cells(monitorRow, 3).Value = wsCleaned.Cells(i, 11).Value
                wsMonitor.Cells(monitorRow, 4).Value = avgDict(isin)
                wsMonitor.Cells(monitorRow, 4).NumberFormat = "0.0000"
                wsMonitor.Cells(monitorRow, 5).Value = deviation
                wsMonitor.Cells(monitorRow, 5).NumberFormat = "0.00%"
                wsMonitor.Cells(monitorRow, 6).Value = "Preisabweichung ueber " & (threshold * 100) & "%"
            End If
        End If
    Next i

    wsMonitor.Columns.AutoFit

End Sub

Sub DelayIssue(wsMonitor As Worksheet, monitorRow As Long, orderID As Variant, isin As Variant, instrumentType As Variant, time As Variant, executionTime As Variant, delay As Long, alert As String)
    monitorRow = monitorRow + 1
    wsMonitor.Cells(monitorRow, 1).Value = orderID
    wsMonitor.Cells(monitorRow, 2).Value = isin
    wsMonitor.Cells(monitorRow, 3).Value = instrumentType
    wsMonitor.Cells(monitorRow, 4).Value = time
    wsMonitor.Cells(monitorRow, 5).Value = executionTime
    wsMonitor.Cells(monitorRow, 6).Value = delay
    wsMonitor.Cells(monitorRow, 7).Value = alert
End Sub

Sub MonitorExecutionDelay()

    On Error Resume Next
    Worksheets("Execution_Delay_Monitoring").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Execution_Delay_Monitoring"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 7)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Order ID"
    wsMonitor.Cells(1, 2).Value = "ISIN"
    wsMonitor.Cells(1, 3).Value = "Security Type"
    wsMonitor.Cells(1, 4).Value = "Timestamp"
    wsMonitor.Cells(1, 5).Value = "Execution Timestamp"
    wsMonitor.Cells(1, 6).Value = "Delay(Seconds)"
    wsMonitor.Cells(1, 7).Value = "Alert"

    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1

    Dim dict As Object
    Dim time As Variant
    Dim isin As Variant
    Dim executionTime As Variant
    Dim status As Variant
    Dim threshold As Double
    Dim delay As Long
    Dim etfDelay As Double
    Dim equityDelay As Double
    Dim bondDelay As Double
    Dim spDelay As Double
    Dim orderID As Variant
    Dim instrumentType As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "ETF", 10
    dict.Add "Equity", 20
    dict.Add "Bond", 60
    dict.Add "Structured Product", 120
    
    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        orderID = wsCleaned.Cells(i, 1).Value
        status = wsCleaned.Cells(i, 8).Value
        time = wsCleaned.Cells(i, 3).Value
        executionTime = wsCleaned.Cells(i, 19).Value
        instrumentType = wsCleaned.Cells(i, 14).Value
        
        If UCase(status) = "EXECUTED" Or UCase(status) = "PARTIALLY EXECUTED" Then
            If Not IsEmpty(time) And Not IsEmpty(executionTime) Then
                delay = DateDiff("s", time, executionTime)
                etfDelay = dict("ETF")
                equityDelay = dict("Equity")
                bondDelay = dict("Bond")
                spDelay = dict("Structured Product")
                If instrumentType = "ETF" And delay > etfDelay Then
                   DelayIssue wsMonitor, monitorRow, orderID, isin, instrumentType, time, executionTime, delay, "Delay > " & etfDelay
                End If
                If instrumentType = "Equity" And delay > equityDelay Then
                   DelayIssue wsMonitor, monitorRow, orderID, isin, instrumentType, time, executionTime, delay, "Delay > " & equityDelay
                End If
                If instrumentType = "Bond" And delay > bondDelay Then
                   DelayIssue wsMonitor, monitorRow, orderID, isin, instrumentType, time, executionTime, delay, "Delay > " & bondDelay
                End If
                If instrumentType = "Structured Product" And delay > spDelay Then
                   DelayIssue wsMonitor, monitorRow, orderID, isin, instrumentType, time, executionTime, delay, "Delay > " & spDelay
                End If
             End If
        End If
            
     Next i
            
    wsMonitor.Columns.AutoFit

End Sub
