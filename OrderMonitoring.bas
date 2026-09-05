Attribute VB_Name = "OrderMonitoring"
Option Explicit
' ============================================================
' Sub: MonitorPendingOrders
' Zweck: Erkennt Orders mit Status "Pending", die laenger als
'        [thresholdMinutes] Minuten unbearbeitet sind.
' Input: Cleaned_Data (Entscheidungsgrundlage: Status, Timestamp;
'        bei Treffer werden zusaetzlich Order ID, ISIN und ClientID
'        fuer die Ergebnisdarstellung ausgelesen)
' Output: Neues Sheet "Pending_Monitoring"
' Bekannte Einschraenkung: Schleife testweise auf Zeile 2-100
'        begrenzt (bewusste Einschraenkung fuer Demozwecke, siehe
'        Kommentar im Code).
' ============================================================
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
' ============================================================
' Sub: MonitorPriceDeviation
' Zweck: Erkennt Ausfuehrungen, deren Preis mehr als [threshold] vom
'        durchschnittlichen Ausfuehrungspreis derselben ISIN abweicht
'        (nur bei Continuous-Trading-Orders; bei der Durchschnitts-
'        berechnung wird die betrachtete Ausfuehrung selbst ausgeschlossen,
'        um eine Verwaesserung des Vergleichswerts durch den eigenen
'        Wert zu vermeiden).
' Input: Cleaned_Data (Spalte ISIN, Execution Price, Eligible_For_Price_Check)
' Output: Neues Sheet "Price_Deviation_Monitoring"
' Bekannte Einschraenkung: Schwellenwert ist fest auf 5% gesetzt und nicht
'        an die tatsaechliche Volatilitaet angepasst - eine dynamische
'        Schwellenwert-Variante ist Teil eines geplanten Python-Moduls.
'        Kommt eine ISIN nur einmal im Datensatz vor, kann der
'        Durchschnitt ohne den eigenen Wert nicht berechnet werden;
'        diese Zeile wird dann uebersprungen.
' ============================================================
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

    Dim dictsum As Object
    Dim dictCount As Object
    Dim isin As Variant
    Dim ep As Variant
    Dim threshold As Double
    Set dictsum = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")

    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        ep = wsCleaned.Cells(i, 11).Value

        If Not IsEmpty(ep) And wsCleaned.Cells(i, 12).Value = "Yes" Then
            If dictsum.Exists(isin) Then
                dictsum(isin) = dictsum(isin) + ep
                dictCount(isin) = dictCount(isin) + 1
            Else
                dictsum(isin) = ep
                dictCount(isin) = 1
            End If
        End If
    Next i

    Dim deviation As Double
    threshold = 0.05
    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        ep = wsCleaned.Cells(i, 11).Value

        If Not IsEmpty(ep) And wsCleaned.Cells(i, 12).Value = "Yes" Then
            If dictCount(isin) > 1 Then
               Dim avgExclSelf As Double
               avgExclSelf = (dictsum(isin) - ep) / (dictCount(isin) - 1)
               deviation = (ep - avgExclSelf) / avgExclSelf

               If Abs(deviation) > threshold Then
                  monitorRow = monitorRow + 1
                  wsMonitor.Cells(monitorRow, 1).Value = wsCleaned.Cells(i, 1).Value
                  wsMonitor.Cells(monitorRow, 2).Value = wsCleaned.Cells(i, 2).Value
                  wsMonitor.Cells(monitorRow, 3).Value = wsCleaned.Cells(i, 11).Value
                  wsMonitor.Cells(monitorRow, 4).Value = avgExclSelf
                  wsMonitor.Cells(monitorRow, 4).NumberFormat = "0.0000"
                  wsMonitor.Cells(monitorRow, 5).Value = deviation
                  wsMonitor.Cells(monitorRow, 5).NumberFormat = "0.00%"
                  wsMonitor.Cells(monitorRow, 6).Value = "Preisabweichung ueber " & (threshold * 100) & "%"
                End If
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
' ============================================================
' Sub: MonitorExecutionDelay
' Zweck: Erkennt Ausfuehrungen, deren Verzoegerung zwischen
'        Auftrags- und Ausfuehrungszeitpunkt den fuer die jeweilige
'        Instrument-Klasse definierten Schwellenwert ueberschreitet
'        (ETF: 10s, Equity: 20s, Bond: 60s, Structured Product: 120s).
' Input: Cleaned_Data (Spalte Timestamp, Execution Time, Instrument Type, Status)
' Output: Neues Sheet "Execution_Delay_Monitoring"
' Bekannte Einschraenkung: Schwellenwerte sind fest und liquiditaets-
'        basiert geschaetzt, nicht empirisch kalibriert. Echte
'        Ausfuehrungsverzoegerungsdaten sind interne Informationen
'        des jeweiligen Instituts und daher nicht zugaenglich.
' ============================================================
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
    
    etfDelay = dict("ETF")
    equityDelay = dict("Equity")
    bondDelay = dict("Bond")
    spDelay = dict("Structured Product")
    
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
' ============================================================
' Sub: MonitorOrderQuantitySpike
' Zweck: Erkennt Orders, deren Menge das [multiple]-fache des
'        durchschnittlichen Ordervolumens desselben Clients
'        (ohne die aktuelle Order) uebersteigt.
' Input: Cleaned_Data (Spalte ClientID, Quantity)
' Output: Neues Sheet "Order_Quantity_Spike_Monitoring"
' Bekannte Einschraenkung: Feste Multiple-Schwelle (3x), keine
'        dynamische/rollierende Baseline - siehe geplantes
'        Python-Modul fuer eine statistisch verfeinerte Variante.
' ============================================================
Sub MonitorOrderQuantitySpike()

    On Error Resume Next
    Worksheets("Order_Quantity_Spike_Monitoring").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Order_Quantity_Spike_Monitoring"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 6)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Order ID"
    wsMonitor.Cells(1, 2).Value = "Client ID"
    wsMonitor.Cells(1, 3).Value = "Quantity"
    wsMonitor.Cells(1, 4).Value = "Avg Quantity (Excl. Current)"
    wsMonitor.Cells(1, 5).Value = "Multiple"
    wsMonitor.Cells(1, 6).Value = "Alert"

    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim orderID As Variant
    Dim clientID As Variant
    Dim quantity As Variant
    Dim monitorRow As Long
    Dim multiple As Integer
    multiple = 3
    monitorRow = 1

    Dim dictsum As Object
    Dim dictCount As Object
    Set dictsum = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")

    For i = 2 To lastCleanRow
        clientID = wsCleaned.Cells(i, 9).Value
        quantity = wsCleaned.Cells(i, 4).Value
        
        If Not IsEmpty(clientID) And Not IsEmpty(quantity) Then
            If dictsum.Exists(clientID) Then
                dictsum(clientID) = dictsum(clientID) + quantity
                dictCount(clientID) = dictCount(clientID) + 1
            Else
                dictsum(clientID) = quantity
                dictCount(clientID) = 1
            End If
        End If
    Next i
    
    Dim avg As Double
    For i = 2 To lastCleanRow
        quantity = wsCleaned.Cells(i, 4).Value
        orderID = wsCleaned.Cells(i, 1).Value
        clientID = wsCleaned.Cells(i, 9).Value

         If Not IsEmpty(clientID) And Not IsEmpty(quantity) Then
             If dictCount(clientID) > 1 Then
                avg = (dictsum(clientID) - quantity) / (dictCount(clientID) - 1)
               If quantity > multiple * avg Then
                  monitorRow = monitorRow + 1
                  wsMonitor.Cells(monitorRow, 1).Value = orderID
                  wsMonitor.Cells(monitorRow, 2).Value = clientID
                  wsMonitor.Cells(monitorRow, 3).Value = quantity
                  wsMonitor.Cells(monitorRow, 4).Value = avg
                  wsMonitor.Cells(monitorRow, 4).NumberFormat = "0.0000"
                  wsMonitor.Cells(monitorRow, 5).Value = multiple
                  wsMonitor.Cells(monitorRow, 6).Value = "Volume exceeds " & multiple & "x average"
                End If
              End If
          End If
     Next i

    wsMonitor.Columns.AutoFit

End Sub



