Attribute VB_Name = "OrderAssetExposure"
Option Explicit

Sub ExposureIssue(wsMonitor As Worksheet, monitorRow As Long, orderID As Variant, isin As Variant, instrumentType As Variant, quantity As Variant, executionPrice As Variant, notionalValue As Double, threshold As Double, alert As String)
    monitorRow = monitorRow + 1
    wsMonitor.Cells(monitorRow, 1).Value = orderID
    wsMonitor.Cells(monitorRow, 2).Value = isin
    wsMonitor.Cells(monitorRow, 3).Value = instrumentType
    wsMonitor.Cells(monitorRow, 4).Value = quantity
    wsMonitor.Cells(monitorRow, 5).Value = executionPrice
    wsMonitor.Cells(monitorRow, 6).Value = notionalValue
    wsMonitor.Cells(monitorRow, 7).Value = threshold
    wsMonitor.Cells(monitorRow, 8).Value = alert
End Sub
' ============================================================
' Sub: MonitorLargeTrade
' Zweck: Erkennt einzelne Ausfuehrungen, deren Notional Value den
'        instrumentklassenspezifischen Grossauftrags-Schwellenwert
'        ueberschreitet.
' Input: Cleaned_Data (Spalte Quantity, Execution Price, Instrument Type, Status)
' Output: Neues Sheet "Large_Trade_Alert"; zusaetzlich werden in
'         Cleaned_Data alle betroffenen Zeilen gelb markiert
'         (Interior.Color = vbYellow) als visueller Hinweis auf
'         Exposure-Ueberschreitungen.
' Bekannte Einschraenkung: Feste Schwellenwerte je Instrument-Klasse
'        (ETF:200.000, Equity:700.000 Bond: 150.000, Structured Product: 50.000),
'        keine dynamische Kalibrierung - die Schwellenwerte haengen von
'        institutsspezifischen Risikoparametern ab, die nicht oeffentlich
'        verfuegbar sind.
' ============================================================
Sub MonitorLargeTrade()

    On Error Resume Next
    Worksheets("Large_Trade_Alert").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Large_Trade_Alert"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 8)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Order ID"
    wsMonitor.Cells(1, 2).Value = "ISIN"
    wsMonitor.Cells(1, 3).Value = "Instrument Type"
    wsMonitor.Cells(1, 4).Value = "Quantity"
    wsMonitor.Cells(1, 5).Value = "Execution Price"
    wsMonitor.Cells(1, 6).Value = "NotionalValue"
    wsMonitor.Cells(1, 7).Value = "Threshold"
    wsMonitor.Cells(1, 8).Value = "Alert"

    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).row

    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1

    Dim dict As Object
    Dim status As Variant
    Dim isin As Variant
    Dim threshold As Double
    Dim quantity As Variant
    Dim orderID As Variant
    Dim instrumentType As Variant
    Dim executionPrice As Variant
    Dim notionalValue As Double
    Dim etfExposure As Long
    Dim equityExposure As Long
    Dim bondExposure As Long
    Dim spExposure As Long
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "ETF", 200000
    dict.Add "Equity", 700000
    dict.Add "Bond", 150000
    dict.Add "Structured Product", 50000
    etfExposure = dict("ETF")
    equityExposure = dict("Equity")
    bondExposure = dict("Bond")
    spExposure = dict("Structured Product")
    
    For i = 2 To lastCleanRow
    
        isin = wsCleaned.Cells(i, 2).Value
        status = wsCleaned.Cells(i, 8).Value
        quantity = wsCleaned.Cells(i, 4).Value
        orderID = wsCleaned.Cells(i, 1).Value
        executionPrice = wsCleaned.Cells(i, 11).Value
        instrumentType = wsCleaned.Cells(i, 14).Value
        
        If Not IsEmpty(quantity) And Not IsEmpty(executionPrice) Then
            If UCase(status) = "EXECUTED" Or UCase(status) = "PARTIALLY EXECUTED" Then
               notionalValue = quantity * executionPrice
                If instrumentType = "ETF" And notionalValue > etfExposure Then
                   wsCleaned.Rows(i).Interior.Color = vbYellow               ' gelb markiert(Interior.Color = vbYellow) als visueller Hinweis auf Exposure-Ueberschreitungen.
                   ExposureIssue wsMonitor, monitorRow, orderID, isin, instrumentType, quantity, executionPrice, notionalValue, dict("ETF"), "Notional Value > " & etfExposure
                End If
                If instrumentType = "Equity" And notionalValue > equityExposure Then
                   wsCleaned.Rows(i).Interior.Color = vbYellow
                   ExposureIssue wsMonitor, monitorRow, orderID, isin, instrumentType, quantity, executionPrice, notionalValue, dict("Equity"), "Notional Value > " & equityExposure
                End If
                If instrumentType = "Bond" And notionalValue / 100 > bondExposure Then
                   wsCleaned.Rows(i).Interior.Color = vbYellow
                   ExposureIssue wsMonitor, monitorRow, orderID, isin, instrumentType, quantity, executionPrice, notionalValue / 100, dict("Bond"), "Notional Value > " & bondExposure
                End If
                If instrumentType = "Structured Product" And notionalValue > spExposure Then
                   wsCleaned.Rows(i).Interior.Color = vbYellow
                   ExposureIssue wsMonitor, monitorRow, orderID, isin, instrumentType, quantity, executionPrice, notionalValue, dict("Structured Product"), "Notional Value > " & spExposure
                End If
             End If
        End If
            
     Next i
            
    wsMonitor.Columns.AutoFit

End Sub

Sub FillExposureRow(wsMonitor As Worksheet, dictBuySum As Object, dictSellSum As Object, i As Long, instrumentType As Variant)
    wsMonitor.Cells(i, 2).Value = dictBuySum(instrumentType)
    wsMonitor.Cells(i, 3).Value = dictSellSum(instrumentType)
    wsMonitor.Cells(i, 4).Value = dictBuySum(instrumentType) - dictSellSum(instrumentType)
    wsMonitor.Cells(i, 5).Value = dictBuySum(instrumentType) + dictSellSum(instrumentType)
End Sub
' ============================================================
' Sub: MonitorAssetExposureInstrumentType
' Zweck: Aggregiert Buy-/Sell-Notional je Instrument-Klasse zu
'        Netto- und Brutto-Exposure ueber den gesamten Datensatz.
' Input: Cleaned_Data (Spalte Quantity, Execution Price, Direction,
'        Instrument Type, Status)
' Output: Neues Sheet "InstrumentTyp_Exposure"
' Bekannte Einschraenkung: keine
' ============================================================
Sub MonitorAssetExposureInstrumentType()

    On Error Resume Next
    Worksheets("InstrumentTyp_Exposure").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "InstrumentTyp_Exposure"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 5)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsMonitor.Cells(1, 1).Value = "Instrument Type"
    wsMonitor.Cells(1, 2).Value = "Buy Notional"
    wsMonitor.Cells(1, 3).Value = "Sell Notional"
    wsMonitor.Cells(1, 4).Value = "Net Exposure"
    wsMonitor.Cells(1, 5).Value = "Gross Exposure"
    
    wsMonitor.Cells(2, 1).Value = "ETF"
    wsMonitor.Cells(3, 1).Value = "Equity"
    wsMonitor.Cells(4, 1).Value = "Bond"
    wsMonitor.Cells(5, 1).Value = "Structured Product"
    
    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).row

    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1
    
    Dim direction As Variant
    Dim quantity As Variant
    Dim status As Variant
    Dim netExposure As Double
    Dim notionalValue As Double
    Dim grossExposure As Double
    Dim instrumentType As Variant
    Dim executionPrice As Variant
    Dim dictBuySum As Object
    Dim dictSellSum As Object
    Set dictBuySum = CreateObject("Scripting.Dictionary")
    Set dictSellSum = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastCleanRow
    
        quantity = wsCleaned.Cells(i, 4).Value
        executionPrice = wsCleaned.Cells(i, 11).Value
        direction = wsCleaned.Cells(i, 5).Value
        instrumentType = wsCleaned.Cells(i, 14).Value
        status = wsCleaned.Cells(i, 8).Value
        
            If Not IsEmpty(quantity) And Not IsEmpty(executionPrice) And Not IsEmpty(instrumentType) Then
                 If UCase(status) = "EXECUTED" Or UCase(status) = "PARTIALLY EXECUTED" Then
                     notionalValue = quantity * executionPrice
                     If instrumentType = "Bond" Then         ' Anleihen werden ueblicherweise als Prozentsatz des Nennwerts notiert(z.B. Kurs 98,5 = 98,5% des Nennwerts), nicht als absoluter Preis pro
                        notionalValue = notionalValue / 100  ' Stueck. Daher muss der Notional Value hier durch 100 geteilt werden,um den tatsaechlichen Gegenwert zu erhalten - bei ETF/Equity/
                     End If                                  ' Structured Product entfaellt diese Umrechnung, da dort der Kurs direkt als Preis pro Einheit notiert wird.
                     If direction = "Buy" Then
                         If dictBuySum.Exists(instrumentType) Then
                            dictBuySum(instrumentType) = dictBuySum(instrumentType) + notionalValue
                         Else
                            dictBuySum(instrumentType) = notionalValue
                         End If
                     Else
                         If dictSellSum.Exists(instrumentType) Then
                            dictSellSum(instrumentType) = dictSellSum(instrumentType) + notionalValue
                         Else
                            dictSellSum(instrumentType) = notionalValue
                         End If
                     End If
                  End If
            End If
            
    Next i
    
    FillExposureRow wsMonitor, dictBuySum, dictSellSum, 2, "ETF"
    FillExposureRow wsMonitor, dictBuySum, dictSellSum, 3, "Equity"
    FillExposureRow wsMonitor, dictBuySum, dictSellSum, 4, "Bond"
    FillExposureRow wsMonitor, dictBuySum, dictSellSum, 5, "Structured Product"
    
    wsMonitor.Columns.AutoFit
    
End Sub
' ============================================================
' Sub: MonitorAssetExposureBySecurity
' Zweck: Aggregiert Buy-/Sell-Notional je ISIN zu Netto- und
'        Brutto-Exposure ueber den gesamten Datensatz.
' Input: Cleaned_Data (Spalte ISIN, Instrument Type, Quantity,
'        Execution Price, Direction, Status)
' Output: Neues Sheet "Security_Exposure"
' Design-Hinweis: Die ISIN-Liste wird bewusst aus Cleaned_Data
'        abgeleitet statt aus Security_Master, damit nur ISINs mit
'        tatsaechlicher Handelsaktivitaet in der Exposure-Uebersicht
'        erscheinen (Security_Master enthaelt auch nie gehandelte
'        ISINs, die hier keinen Mehrwert haetten).
' Bekannte Einschraenkung: keine
' ============================================================
Sub MonitorAssetExposureBySecurity()

    On Error Resume Next
    Worksheets("Security_Exposure").Delete
    On Error GoTo 0

    Dim wsMonitor As Worksheet
    Set wsMonitor = Worksheets.Add
    wsMonitor.Name = "Security_Exposure"

    Dim wsCleaned As Worksheet
    Set wsCleaned = Worksheets("Cleaned_Data")

    wsCleaned.Cells(1, 11).Copy
    wsMonitor.Range(wsMonitor.Cells(1, 1), wsMonitor.Cells(1, 6)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    wsMonitor.Cells(1, 1).Value = "ISIN"
    wsMonitor.Cells(1, 2).Value = "Instrument Type"
    wsMonitor.Cells(1, 3).Value = "Buy Notional"
    wsMonitor.Cells(1, 4).Value = "Sell Notional"
    wsMonitor.Cells(1, 5).Value = "Net Exposure"
    wsMonitor.Cells(1, 6).Value = "Gross Exposure"
    
    Dim lastCleanRow As Long
    lastCleanRow = wsCleaned.Cells(wsCleaned.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    Dim monitorRow As Long
    monitorRow = 1
    
    Dim isin As Variant
    Dim direction As Variant
    Dim quantity As Variant
    Dim status As Variant
    Dim netExposure As Double
    Dim notionalValue As Double
    Dim grossExposure As Double
    Dim instrumentType As Variant
    Dim executionPrice As Variant
    Dim dictBuySum As Object
    Dim dictSellSum As Object
    Dim dictISIN As Object
    Set dictBuySum = CreateObject("Scripting.Dictionary")
    Set dictSellSum = CreateObject("Scripting.Dictionary")
    Set dictISIN = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastCleanRow
        isin = wsCleaned.Cells(i, 2).Value
        instrumentType = wsCleaned.Cells(i, 14).Value
        If Not IsEmpty(isin) And Not dictISIN.Exists(isin) And Not IsEmpty(instrumentType) Then
           monitorRow = monitorRow + 1
           wsMonitor.Cells(monitorRow, 1).Value = isin
           wsMonitor.Cells(monitorRow, 2).Value = instrumentType
           dictISIN(isin) = 0                    ' Diese erste Schleife ermittelt zunaechst alle eindeutigen ISINs und
        End If                                   ' legt die Zeilenstruktur im Sheet fest. Im Gegensatz zu
    Next i                                       ' MonitorAssetExposureInstrumentType (wo die 4 Instrument-Klassen fest
                                                 ' vorgegeben sind) ist die Anzahl und Reihenfolge der ISINs hier vorher
    For i = 2 To lastCleanRow                    ' nicht bekannt - daher muss diese Liste erst aus den Daten ermittelt
        quantity = wsCleaned.Cells(i, 4).Value   ' werden, bevor die zweite Schleife die Buy-/Sell-Betraege pro ISIN
        executionPrice = wsCleaned.Cells(i, 11).Value ' zuordnen kann.
        direction = wsCleaned.Cells(i, 5).Value
        status = wsCleaned.Cells(i, 8).Value
        isin = wsCleaned.Cells(i, 2).Value
        instrumentType = wsCleaned.Cells(i, 14).Value
            If Not IsEmpty(quantity) And Not IsEmpty(executionPrice) And Not IsEmpty(isin) Then
                 If UCase(status) = "EXECUTED" Or UCase(status) = "PARTIALLY EXECUTED" Then
                     notionalValue = quantity * executionPrice
                     If instrumentType = "Bond" Then
                        notionalValue = notionalValue / 100
                     End If
                     If direction = "Buy" Then
                         If dictBuySum.Exists(isin) Then
                            dictBuySum(isin) = dictBuySum(isin) + notionalValue
                         Else
                            dictBuySum(isin) = notionalValue
                         End If
                     Else
                         If dictSellSum.Exists(isin) Then
                            dictSellSum(isin) = dictSellSum(isin) + notionalValue
                         Else
                            dictSellSum(isin) = notionalValue
                         End If
                     End If
                  End If
            End If
     Next i
     
     Dim lastRow As Long
     lastRow = wsMonitor.Cells(wsMonitor.Rows.Count, 1).End(xlUp).row

     For i = 2 To lastRow
         isin = wsMonitor.Cells(i, 1).Value
         wsMonitor.Cells(i, 3).Value = dictBuySum(isin)
         wsMonitor.Cells(i, 4).Value = dictSellSum(isin)
         wsMonitor.Cells(i, 5).Value = dictBuySum(isin) - dictSellSum(isin)
         wsMonitor.Cells(i, 6).Value = dictBuySum(isin) + dictSellSum(isin)
    Next i
    
    wsMonitor.Columns.AutoFit
    
End Sub

