Attribute VB_Name = "OrderOutput"
Option Explicit

' Liest InstrumentTyp_Exposure (Layer3) und summiert die Gross-Exposure-Spalte
' ueber die vier Asset-Klassen-Zeilen. Keine Neuberechnung - reine Uebernahme
' des dort bereits ermittelten Ergebnisses.
Function GetTotalGrossExposure() As Double
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets("InstrumentTyp_Exposure")
    On Error GoTo 0

    If ws Is Nothing Then
        GetTotalGrossExposure = 0
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim total As Double
    Dim i As Long
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, 5).Value) Then
            total = total + ws.Cells(i, 5).Value
        End If
    Next i

    GetTotalGrossExposure = total
End Function

' Zaehlt die Anzahl der Datenzeilen im angegebenen Sheet (ohne Kopfzeile)
Function CountMonitoringRows(sheetName As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then        ' Falls das Sheet nicht existiert oder leer ist, wird 0 zurueckgegeben,
        CountMonitoringRows = 0  ' um Fehler in nachgelagerten Modulen durch Zugriff auf ein unguelltiges Objekt zu vermeiden
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    If lastRow < 2 Then          ' Nur Kopfzeile vorhanden oder Sheet komplett leer
        CountMonitoringRows = 0
    Else
        CountMonitoringRows = lastRow - 1  ' Kopfzeile abziehen, um die tatsaetchliche Datenzeilenanzahl zu erhalten
    End If
End Function

' Fuegt alle eindeutigen Order-IDs aus dem angegebenen Sheet in das uebergebene Set (Dictionary) ein
' Wird genutzt, um Ergebnisse aus mehreren Monitoring-Sheets in einem gemeinsamen Set zu buendeln und Duplikate zu vermeiden
' Das Dictionary-Objekt wird extern erstellt und per Referenz uebergeben, damit mehrere Aufrufe dasselbe Set befuellen koennen
Sub AddOrderIDsToSet(targetSet As Object, sheetName As String, orderIDCol As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim i As Long
    Dim orderID As Variant
    For i = 2 To lastRow
        orderID = ws.Cells(i, orderIDCol).Value
      If Not IsEmpty(orderID) And orderID <> "" Then  ' Nur hinzufuegen, wenn noch nicht vorhanden (Duplikate vermeiden)
          If Not targetSet.Exists(orderID) Then
             targetSet(orderID) = True
          End If
      End If
    Next i
End Sub

' Ermittelt die Top-N Eintraege aus einem Dictionary, sortiert nach Value absteigend
' Wird genutzt, um z.B. die groessten Kundenengagements oder Positionen zu identifizieren (Konzentrationsrisiko)
' Rueckgabe erfolgt ueber ByRef-Parameter (outKeys, outValues), da eine Function nur einen Wert zurueckgeben kann,
' hier aber zwei zusammengehoerige Arrays (Keys und Values) benoetigt werden
Sub GetTopN(sourceDict As Object, n As Long, ByRef outKeys() As Variant, ByRef outValues() As Double)

    Dim cnt As Long
    cnt = sourceDict.Count   ' Anzahl der Eintraege im Dictionary

    Dim keys() As Variant
    Dim values() As Double
    Dim m As Long
    m = 0

    If cnt > 0 Then           ' Dictionary in zwei parallele Arrays umwandeln (Keys und Values an gleicher Indexposition),
        ReDim keys(cnt - 1)   ' da Dictionary selbst keine Sortierung nach Value unterstuetzt
        ReDim values(cnt - 1)
        Dim k As Variant
        For Each k In sourceDict.keys
            keys(m) = k
            values(m) = sourceDict(k)
            m = m + 1
        Next k
    End If

    Dim resultCount As Long
    resultCount = WorksheetFunction.Min(n, cnt)   ' Tatsaechliche Anzahl der Rueckgabewerte: Minimum aus gewuenschtem N und vorhandener Anzahl
    If resultCount < 0 Then resultCount = 0

    ReDim outKeys(IIf(resultCount > 0, resultCount - 1, 0))
    ReDim outValues(IIf(resultCount > 0, resultCount - 1, 0))

    Dim i As Long, j As Long, maxIdx As Long
    Dim tempKey As Variant
    Dim tempVal As Double

    For i = 0 To cnt - 2
        maxIdx = i    ' Startannahme, wird im inneren Loop ueberprueft und ggf. korrigiert

        For j = i + 1 To cnt - 1
            If values(j) > values(maxIdx) Then maxIdx = j   ' Selection Sort: In jeder Runde wird aus dem unsortierten Rest der groesste Wert gesucht
        Next j                                              ' und an die Position (i) getauscht.
        If maxIdx <> i Then
            tempVal = values(i): values(i) = values(maxIdx): values(maxIdx) = tempVal
            tempKey = keys(i): keys(i) = keys(maxIdx): keys(maxIdx) = tempKey         'Keys und Values werden synchron getauscht,damit die Zuordnung Kunde/Wert erhalten bleibt.
        End If
    Next i

    If resultCount = 0 Then           ' Ergebnis in die Ausgabeparameter kopieren (nur die ersten resultCount, da bereits absteigend sortiert)
        outKeys(0) = "(keine Daten)"  ' Hinweis: Dieser Zweig deckt auch den Fall eines leeren Dictionarys ab (cnt=0), da resultCount in diesem Fall ebenfalls 0 ist
        outValues(0) = 0
    Else
        For i = 0 To resultCount - 1
            outKeys(i) = keys(i)
            outValues(i) = values(i)
        Next i
    End If

End Sub

' ============================================================
' Sub: BuildSummary
' Zweck: Erstellt ein "Summary"-Sheet mit den Kennzahlen, die ein
'        Trader/Supervisor auf den ersten Blick sehen will:
'        Anzahl auffaelliger Orders je Kategorie, konzernweites
'        Gross Exposure (aus Layer3 uebernommen), Top-3-Kunden nach
'        betroffenem Volumen, und die Auffaelligkeitsrate je
'        Asset-Klasse (Anzahl auffaellig / Gesamtanzahl).
' Input: Price_Deviation_Monitoring, Execution_Delay_Monitoring,
'        Order_Quantity_Spike_Monitoring, Large_Trade_Alert
'        (fuer Kategorie-Zaehlung), Cleaned_Data (fuer ClientID,
'        InstrumentType, Quantity, Execution Price je OrderID und
'        fuer die Gesamtanzahl Orders je Asset-Klasse),
'        InstrumentTyp_Exposure (fuer Gross Exposure, aus Layer3
'        uebernommen statt neu berechnet)
' Output: Neues Sheet "Summary"
' Design-Hinweis: Das Gesamt-Exposure wird bewusst NICHT aus den vier
'        Anomalie-Kategorien zusammengerechnet (das waere eine
'        Vermischung von "auffaelligem Volumen" und "Gesamt-Exposure"
'        ohne klare Aussage) - stattdessen direkte Uebernahme des
'        bereits in Layer3 berechneten Gross Exposure je Asset-Klasse.
'        Die Asset-Klassen-Verteilung zeigt eine Rate (auffaellig /
'        gesamt), nicht nur eine rohe Anzahl, da eine rohe Anzahl ohne
'        Nenner keine Aussage ueber die tatsaechliche Auffaelligkeit
'        einer Asset-Klasse zulaesst.
' ============================================================
Sub BuildSummary()

  On Error Resume Next
  Worksheets("Summary").Delete
  On Error GoTo 0
  
  Dim wsSummary As Worksheet
  Set wsSummary = Worksheets.Add
  wsSummary.Name = "Summary"

  Dim wsCleaned As Worksheet
  Set wsCleaned = Worksheets("Cleaned_Data")
  
  ' --- Schritt 1: Kategorie-Zaehlung direkt aus den vier Monitoring-Sheets ---
  Dim countPriceDeviation As Long
  Dim countExecutionDelay As Long
  Dim countOrderSpike As Long
  Dim countLargeTrade As Long
  
  countPriceDeviation = CountMonitoringRows("Price_Deviation_Monitoring")    'vier Anomalie-Kategorien
  countExecutionDelay = CountMonitoringRows("Execution_Delay_Monitoring")
  countOrderSpike = CountMonitoringRows("Order_Quantity_Spike_Monitoring")
  countLargeTrade = CountMonitoringRows("Large_Trade_Alert")
  
  ' --- Schritt 2: Vereinigungsmenge aller auffaelligen OrderIDs bilden ---
  Dim flaggedOrders As Object
  Set flaggedOrders = CreateObject("Scripting.Dictionary")
  
  AddOrderIDsToSet flaggedOrders, "Price_Deviation_Monitoring", 1
  AddOrderIDsToSet flaggedOrders, "Execution_Delay_Monitoring", 1
  AddOrderIDsToSet flaggedOrders, "Order_Quantity_Spike_Monitoring", 1
  AddOrderIDsToSet flaggedOrders, "Large_Trade_Alert", 1
  
  ' --- Schritt 3: OrderID -> Zeilennummer in Cleaned_Data einmalig aufbauen,
  '     gleichzeitig Gesamtanzahl Orders je Asset-Klasse zaehlen (Nenner fuer die Rate) ---
  Dim orderRowMap As Object
  Set orderRowMap = CreateObject("Scripting.Dictionary")
  
  Dim totalCountByType As Object
  Set totalCountByType = CreateObject("Scripting.Dictionary")

  Dim lastCleanRow As Long
  lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).row
  
  Dim i As Long
  Dim orderID As Variant
  Dim rowInstrumentType As Variant
  
  For i = 2 To lastCleanRow
        orderID = wsCleaned.Cells(i, 1).Value
        If Not IsEmpty(orderID) And orderID <> "" Then  ' Sicherheitspruefung: Die Daten sollten nach der Data-Cleaning-Schicht bereits vollstaendig sein
            If Not orderRowMap.Exists(orderID) Then     ' Diese Pruefung dient als zusaetzliche Absicherung gegen unerwartete Luecken, kein Hinweis auf einen Fehler im Cleaning-Prozess
                orderRowMap(orderID) = i
            End If
        End If
        
        rowInstrumentType = wsCleaned.Cells(i, 14).Value
        If Not IsEmpty(rowInstrumentType) And rowInstrumentType <> "" Then
            If totalCountByType.Exists(rowInstrumentType) Then                  'Bereinigtes Sheet nach Instrument Type gruppiert und gezaehlt
                totalCountByType(rowInstrumentType) = totalCountByType(rowInstrumentType) + 1
            Else
                totalCountByType(rowInstrumentType) = 1
            End If
        End If
  Next i
  
  ' --- Schritt 4: fuer jede auffaellige OrderID Kunde (nach Volumen) und
  '     Asset-Klasse (Anzahl auffaellig) akkumulieren ---
  Dim clientAmountDict As Object
  Dim assetCountDict As Object
  Set clientAmountDict = CreateObject("Scripting.Dictionary")
  Set assetCountDict = CreateObject("Scripting.Dictionary")

  Dim row As Long
  Dim quantity As Variant
  Dim execPrice As Variant
  Dim clientID As Variant
  Dim instrumentType As Variant
  Dim notionalValue As Double
  
  For Each orderID In flaggedOrders.keys
    If orderRowMap.Exists(orderID) Then
        row = orderRowMap(orderID)
        quantity = wsCleaned.Cells(row, 4).Value
        execPrice = wsCleaned.Cells(row, 11).Value
        clientID = wsCleaned.Cells(row, 9).Value
        instrumentType = wsCleaned.Cells(row, 14).Value

         If Not IsEmpty(instrumentType) And instrumentType <> "" Then
                If assetCountDict.Exists(instrumentType) Then          'Zaehlung von flaggedOrders (ein aus vier Anomalie-Sheets zusammengesetztes Set), gruppiert nach Instrument Type
                    assetCountDict(instrumentType) = assetCountDict(instrumentType) + 1
                Else
                    assetCountDict(instrumentType) = 1
                End If
         End If

         If Not IsEmpty(quantity) And Not IsEmpty(execPrice) And quantity <> "" And execPrice <> "" Then
                notionalValue = quantity * execPrice
                If instrumentType = "Bond" Then
                    notionalValue = notionalValue / 100
                End If

                If Not IsEmpty(clientID) And clientID <> "" Then
                    If clientAmountDict.Exists(clientID) Then
                        clientAmountDict(clientID) = clientAmountDict(clientID) + notionalValue
                    Else
                        clientAmountDict(clientID) = notionalValue
                   End If
                End If
         End If
     End If
  Next orderID

  ' --- Schritt 5: Gross Exposure direkt aus Layer3 (InstrumentTyp_Exposure) uebernehmen ---
  Dim totalGrossExposure As Double
  totalGrossExposure = GetTotalGrossExposure()
  
  ' --- Schritt 6: Ergebnisse in feste Zellen schreiben ---
  wsSummary.Cells(1, 1).Value = "Daily Anomaly Summary"
  wsSummary.Cells(1, 1).Font.Bold = True
  wsSummary.Cells(1, 1).Font.Size = 14
  
  wsCleaned.Cells(1, 11).Copy
  wsSummary.Range(wsSummary.Cells(3, 1), wsSummary.Cells(7, 1)).PasteSpecial Paste:=xlPasteFormats
  Application.CutCopyMode = False
  
  wsSummary.Cells(3, 1).Value = "Auffaellige Orders gesamt"
  wsSummary.Cells(3, 2).Value = flaggedOrders.Count
  wsSummary.Cells(3, 4).Value = "(dedupliziert; Summe der Einzeltypen kann hoeher sein, da Auftraege mehrere Auffaelligkeiten aufweisen koennen)"
  wsSummary.Cells(4, 1).Value = "  davon Preisabweichung"
  wsSummary.Cells(4, 2).Value = countPriceDeviation
  wsSummary.Cells(5, 1).Value = "  davon Ausfuehrungsverzoegerung"
  wsSummary.Cells(5, 2).Value = countExecutionDelay
  wsSummary.Cells(6, 1).Value = "  davon Order-Mengen-Spike"
  wsSummary.Cells(6, 2).Value = countOrderSpike
  wsSummary.Cells(7, 1).Value = "  davon Exposure-Limit ueberschritten"
  wsSummary.Cells(7, 2).Value = countLargeTrade
  
  wsSummary.Cells(9, 1).Value = "Konzernweites Gross Exposure"
  wsSummary.Cells(9, 2).Value = totalGrossExposure
  wsSummary.Cells(9, 2).NumberFormat = "#,##0.00"

  wsSummary.Cells(11, 1).Value = "Top 3 auffaellige Kunden (nach Volumen)"
  wsSummary.Cells(11, 1).Font.Bold = True
  
   Dim topClients() As Variant
   Dim topAmounts() As Double
   GetTopN clientAmountDict, 3, topClients, topAmounts

   Dim r As Long
   For r = 0 To UBound(topClients) 'um zu verhindern, dass bei ungewoehnlich wenigen Kunden an einem bestimmten Tag ein Array-Index-Fehler (Out of Bounds) auftritt
        wsSummary.Cells(12 + r, 1).Value = topClients(r)
        wsSummary.Cells(12 + r, 2).Value = topAmounts(r)
        wsSummary.Cells(12 + r, 2).NumberFormat = "#,##0.00"
   Next r

   wsSummary.Cells(16, 1).Value = "Auffaelligkeitsrate nach Asset-Klasse"
   wsSummary.Cells(16, 1).Font.Bold = True
   wsSummary.Cells(16, 2).Value = "Auffaellig"
   wsSummary.Cells(16, 3).Value = "Gesamt"
   wsSummary.Cells(16, 4).Value = "Rate"

   Dim assetTypes As Variant
        assetTypes = Array("Equity", "ETF", "Bond", "Structured Product")

   Dim a As Long
   Dim anomalyCnt As Long
   Dim totalCnt As Long

   For a = 0 To UBound(assetTypes)  'So geschrieben, lassen sich jederzeit weitere Asset-Klassen hinzufuegen, ohne den Loop-Code anzufassen.
        wsSummary.Cells(17 + a, 1).Value = assetTypes(a)

        If assetCountDict.Exists(assetTypes(a)) Then
            anomalyCnt = assetCountDict(assetTypes(a))
        Else
            anomalyCnt = 0
        End If

        If totalCountByType.Exists(assetTypes(a)) Then
            totalCnt = totalCountByType(assetTypes(a))
        Else
            totalCnt = 0
        End If

        wsSummary.Cells(17 + a, 2).Value = anomalyCnt
        wsSummary.Cells(17 + a, 3).Value = totalCnt

        If totalCnt > 0 Then
            wsSummary.Cells(17 + a, 4).Value = anomalyCnt / totalCnt
            wsSummary.Cells(17 + a, 4).NumberFormat = "0.00%"
        Else
            wsSummary.Cells(17 + a, 4).Value = "n/a"  'totalCnt = 0
        End If
   Next a
  
  wsSummary.Columns.AutoFit
    
End Sub
