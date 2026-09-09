Attribute VB_Name = "OrderDataCleaning"
Option Explicit
'============================================================
' Layer¢Ù: Data Cleaning
' Zweck: Bereinigt Raw_Client_Orders zeilenweise: validiert Format
'        und Referenzintegritaet (ISIN gegen Security_Master,
'        ClientID gegen Client_Master), parst Preise/Mengen/
'        Zeitstempel und prueft feld¨¹bergreifende Konsistenz
'        (z.B. Limit- vs. Execution Price, Status vs. Ausfuehrungsdate
'        Zeitstempel-Reihenfolge, doppelte OrderIDs).
' Input: Raw_Client_Orders, Security_Master, Client_Master
' Output: Neue Sheets "Cleaned_Data" und "Data_Quality"
' Bekannte Einschraenkung: Der normale Arbeitsablauf sieht vor, dass
'        nach Abschluss der Datenbereinigung anhand der Tabelle
'        Data_Quality alle verbliebenen Luecken in Cleaned_Data
'        (Rueckstaende, unklare Faelle etc.) behoben werden, bevor
'        mit Layer¢Ú/¢Û fortgefahren wird - andernfalls beeintraechtigen
'        die verbliebenen Luecken zwangslaeufig in gewissem Umfang
'        die Ergebnisse der beiden nachfolgenden Ebenen. Dieser manuelle
'        Schritt wird im vorliegenden Projekt (noch) nicht abgebildet.
' ============================================================

Sub LogQualityIssue(wsQuality As Worksheet, ByRef qualityRow As Long, rowNum As Long, orderIDValue As Variant, problemDesc As String)
    qualityRow = qualityRow + 1
    wsQuality.Cells(qualityRow, 1).Value = rowNum
    wsQuality.Cells(qualityRow, 2).Value = orderIDValue
    wsQuality.Cells(qualityRow, 3).Value = problemDesc
End Sub

Function ParsePrice(txt As Variant) As Variant
    If IsError(txt) Then
        ParsePrice = Null
        Exit Function
    End If

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

    If InStr(cleanTxt, "€") > 0 Then
        cleanTxt = Replace(cleanTxt, "€", "")
    End If

    Dim lastComma As Long
    Dim lastDot As Long
    lastComma = InStrRev(cleanTxt, ",")             ' Erkennung: das Zeichen (Komma oder Punkt), das am weitesten hinten steht,
    lastDot = InStrRev(cleanTxt, ".")               '            ist das Dezimaltrennzeichen; das davor liegende ist der Tausendertrenner.
                                                    ' Beispiel: "1.234,56" (EU) -> Komma hinten -> Dezimaltrennzeichen = Komma
   If lastComma > 0 And lastDot > 0 Then            '           "1,234.56" (US) -> Punkt hinten -> Dezimaltrennzeichen = Punkt
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

Sub DataCleaning()
    
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
    wsCleaned.Range(wsCleaned.Cells(1, 12), wsCleaned.Cells(1, 19)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    wsCleaned.Cells(1, 12).Value = "Eligible_For_Price_Check"
    wsCleaned.Cells(1, 13).Value = "Security Name"
    wsCleaned.Cells(1, 14).Value = "Instrument Type"
    wsCleaned.Cells(1, 15).Value = "Currency"
    wsCleaned.Cells(1, 16).Value = "Client Name"
    wsCleaned.Cells(1, 17).Value = "Client Type"
    wsCleaned.Cells(1, 18).Value = "Region"
    wsCleaned.Cells(1, 19).Value = "Execution Time"

    wsQuality.Cells(1, 1).Value = "Row"
    wsQuality.Cells(1, 2).Value = "OrderID"
    wsQuality.Cells(1, 3).Value = "Problem"
    Dim col As Long
    For col = 1 To 20
    wsCleaned.Cells(1, col).AddComment "Hinweis: Gelb markierte Zeilen kennzeichnen Auftraege, deren Notional Value den definierten Grossauftrags-Schwellenwert ueberschreitet " & _
                                                "(Details zu Notional Value und Schwellenwert je Auftrag siehe Large_Trade_Alert sowie Sub MonitorLargeTrade)."
    wsCleaned.Cells(1, col).Comment.Shape.TextFrame.AutoSize = True
    With wsCleaned.Cells(1, col).Comment.Shape
        .TextFrame.AutoSize = False
        .Width = 200
        .Height = 100
        .TextFrame.Characters.Font.Size = 11
    End With
    Next col

    Dim wsData As Worksheet
    Dim wsSMData As Worksheet
    Dim wsCMData As Worksheet

    Set wsData = Worksheets("Raw_Client_Orders")
    Set wsSMData = Worksheets("Security_Master")
    Set wsCMData = Worksheets("Client_Master")

    Dim i As Long
    Dim lastRow As Long
    Dim qualityRow As Long
    Dim rowNum As Long

    Dim parsed As Variant
    Dim parsedExecuTime As Variant
    Dim parsedLimitPrice As Variant
    Dim parsedExecutionPrice As Variant
    Dim parsedQty As Variant

    Dim isISINValid As Boolean
    Dim isISINFound As Boolean
    Dim isClientIDValid As Boolean
    Dim isClientIDFound As Boolean
    Dim isOrderIDValid As Boolean
    Dim isMarketPhaseValid As Boolean

    Dim orderID As String
    Dim cleanedOrderID As String
    Dim rawISIN As String
    Dim cleanedISIN As String
    Dim rawStatus As String
    Dim rawClientID As String
    Dim cleanedClientID As String
    Dim orderType As String
    Dim rawExceuTime As Variant

    Dim dict As Object
    Dim key As Variant
    Dim rowList As String

    Set dict = CreateObject("Scripting.Dictionary")

    qualityRow = 1
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    Dim secDict As Object
    Set secDict = CreateObject("Scripting.Dictionary")

    Dim lastSecRow As Long
    lastSecRow = wsSMData.Cells(wsSMData.Rows.Count, 1).End(xlUp).row
                             
    Dim j As Long            ' Hinweis: Security_Master und Client_Master werden als bereits bereinigte, vertrauenswuerdige Referenzdaten behandelt (im Gegensatz zu
                             ' Raw_Client_Orders) - daher hier bewusst kein Trim beim Aufbau des Dictionary.
    For j = 2 To lastSecRow
        secDict(wsSMData.Cells(j, 1).Value) = Array(wsSMData.Cells(j, 2).Value, wsSMData.Cells(j, 3).Value, wsSMData.Cells(j, 4).Value)
    Next j

    Dim cliDict As Object
    Set cliDict = CreateObject("Scripting.Dictionary")

    Dim lastCliRow As Long
    lastCliRow = wsCMData.Cells(wsCMData.Rows.Count, 1).End(xlUp).row

    Dim k As Long
    For k = 2 To lastCliRow
        cliDict(wsCMData.Cells(k, 1).Value) = Array(wsCMData.Cells(k, 2).Value, wsCMData.Cells(k, 3).Value, wsCMData.Cells(k, 4).Value)
    Next k

    For i = 2 To lastRow

        orderID = wsData.Cells(i, 1).Value
        cleanedOrderID = Trim(CStr(orderID))
        parsed = ParseTimestamp(wsData.Cells(i, 3).Value)
        parsedExecuTime = ParseTimestamp(wsData.Cells(i, 12).Value)
        parsedLimitPrice = ParsePrice(wsData.Cells(i, 7).Value)
        parsedExecutionPrice = ParsePrice(wsData.Cells(i, 11).Value)
        parsedQty = ParseQuantity(wsData.Cells(i, 4).Value)
        isMarketPhaseValid = ValidateMarketPhaseFormat(wsData.Cells(i, 10).Value)
        isOrderIDValid = ValidateOrderIDFormat(wsData.Cells(i, 1).Value)

        rawISIN = wsData.Cells(i, 2).Value
        cleanedISIN = Trim(rawISIN)
        rawStatus = wsData.Cells(i, 8).Value
        rawClientID = wsData.Cells(i, 9).Value
        rawExceuTime = wsData.Cells(i, 12).Value
        cleanedClientID = Trim(rawClientID)

        isISINValid = ValidateISINFormat(rawISIN)
        If Not isISINValid Then                                                                                           ' Format ist korrekt, aber ISIN nicht in Referenzdatenbank gefunden -
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "ISIN is empty or invalid ISIN format"    ' unterscheidet sich von Faellen mit ungueltigem ISIN-Format: obwohl
        Else                                                                                                              ' in Cleaned_Data beide Faelle nicht erscheinen, sind die Ursachen
            isISINFound = secDict.Exists(cleanedISIN)                                                                     ' unterschiedlich (Formatfehler vs. fehlender Referenzeintrag).
            If Not isISINFound Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "ISIN not found in Security_Master"
            Else
                wsCleaned.Cells(i, 2).Value = cleanedISIN
                wsCleaned.Cells(i, 13).Value = secDict(cleanedISIN)(0)
                wsCleaned.Cells(i, 14).Value = secDict(cleanedISIN)(1)
                wsCleaned.Cells(i, 15).Value = secDict(cleanedISIN)(2)
            End If
        End If

        isClientIDValid = ValidateClientIDFormat(rawClientID)
        If Not isClientIDValid Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "ClientID is empty or invalid ClientID format"
        Else
            isClientIDFound = cliDict.Exists(cleanedClientID)
            If Not isClientIDFound Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "ClientID not found in Client_Master"
            Else
                wsCleaned.Cells(i, 9).Value = cleanedClientID
                wsCleaned.Cells(i, 16).Value = cliDict(cleanedClientID)(0)
                wsCleaned.Cells(i, 17).Value = cliDict(cleanedClientID)(1)
                wsCleaned.Cells(i, 18).Value = cliDict(cleanedClientID)(2)
            End If
        End If

        If Not isMarketPhaseValid Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "MarketPhase is empty or invalid"
        Else
            wsCleaned.Cells(i, 10).Value = wsData.Cells(i, 10).Value
            If UCase(Trim(wsData.Cells(i, 10).Value)) = "CONTINUOUS" Then
                wsCleaned.Cells(i, 12).Value = "Yes"
            Else
                wsCleaned.Cells(i, 12).Value = "No"
            End If
        End If

        If Not ValidateStatusFormat(rawStatus) Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Status is empty or invalid"
        Else
            wsCleaned.Cells(i, 8).Value = wsData.Cells(i, 8).Value
        End If

        If IsNull(parsedQty) Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Quantity is empty or not a number"
        ElseIf parsedQty < 0 Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Negative Quantity"  ' Negative Menge wird nur protokolliert, nicht automatisch korrigiert -
        Else                                                                                         ' Vorzeichenfehler koennen mehrere Ursachen haben (Sell/Buy vertauscht,
            wsCleaned.Cells(i, 4).Value = parsedQty                                                  ' Tippfehler), daher manuelle Pruefung ueber Data_Quality erforderlich.
        End If

        If IsNull(parsed) Then
            If IsEmpty(wsData.Cells(i, 3).Value) Or wsData.Cells(i, 3).Value = "" Or Trim(CStr(wsData.Cells(i, 3).Value)) = "" Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Timestamp is empty"
            Else
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Achtung: Unbekanntes Datenformat"
            End If
        Else
            wsCleaned.Cells(i, 3).Value = parsed
            wsCleaned.Cells(i, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If

        orderType = Trim(CStr(wsData.Cells(i, 6).Value))

        If UCase(orderType) = "LIMIT ORDER" Then
            If IsNull(parsedLimitPrice) Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "LimitPrice is empty or invalid"
            ElseIf parsedLimitPrice = 0 Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "In der Spalte Limit Price koennten Parsing-Fehler auftreten" ' Val() gibt bei nicht erkennbarem Text 0 zurueck (nicht Null/Fehler) -
            Else                                                                                                                                  ' daher ist 0 hier vermutlich kein echter Preis, sondern ein durch Val()
                wsCleaned.Cells(i, 7).Value = parsedLimitPrice                                                                                    ' verursachter stiller Parsing-Fehler. Manuelle Pruefung der Rohdaten
            End If                                                                                                                                ' erforderlich, um die genaue Ursache zu identifizieren.
        ElseIf UCase(orderType) = "MARKET ORDER" Then
            If Not IsNull(parsedLimitPrice) Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Order Type und Limit Price widerspruechlich - manuelle Pruefung erforderlich"
            End If
        End If

        If UCase(Trim(rawStatus)) = "EXECUTED" Or UCase(Trim(rawStatus)) = "PARTIALLY EXECUTED" Then
            If IsNull(parsedExecutionPrice) Or IsEmpty(rawExceuTime) Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Executed but Execution Price or(and) Executon Time is(are) empty oder status is false"
            Else
                wsCleaned.Cells(i, 11).Value = parsedExecutionPrice
                wsCleaned.Cells(i, 19).Value = parsedExecuTime
            End If
        Else
            If Not IsNull(parsedExecutionPrice) Or Not IsEmpty(rawExceuTime) Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Ausfuerungspreis oder(und) Ausfuerungszeit vorhanden, Status vermutlich veraltet"
                wsCleaned.Cells(i, 11).Value = parsedExecutionPrice ' Status hat bereits die Formatpruefung bestanden und wurde in Cleaned_Data uebernommen - diese Pruefung ist ein zweiter, unabhaengiger Schritt: ein
                wsCleaned.Cells(i, 19).Value = parsedExecuTime      ' formal gueltiger Status (z.B. "Pending") kann trotzdem im Widerspruch zu anderen Feldern derselben Zeile stehen (hier: Execution-Daten vorhanden,
                wsCleaned.Cells(i, 8).Value = Null                  ' obwohl Status "nicht ausgefuehrt" sagt). Die Execution-Daten werden uebernommen(sie sind plausibel vorhanden£¬und manuelle Pruefung erforderlich,
            End If                                                  ' ob der Status haengen geblieben ist), der widerspruechliche Status wird bewusst auf Null gesetzt,um Fehlinterpretation in Layer¢Ú/¢Û zu vermeiden.
        End If

        If Not isOrderIDValid Then
            LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "OrderID is empty or invalid"
        Else
            If dict.Exists(cleanedOrderID) Then
                dict(cleanedOrderID) = dict(cleanedOrderID) & "," & i
            Else
                dict(cleanedOrderID) = i
            End If
        End If

        If Not IsNull(parsed) And Not IsNull(parsedExecuTime) Then
            If parsedExecuTime < parsed Then
                LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Ausfuehrungszeit liegt vor Auftragszeit (Zeitstempel-Konflikt)"
                wsCleaned.Cells(i, 3).Value = Null      ' Auftragszeit liegt nach Ausfuehrungszeit - logisch unmoeglich, aber unklar welcher Zeitstempel falsch ist. Beide werden genullt und ueber
                wsCleaned.Cells(i, 19).Value = Null     ' Data_Quality zur gemeinsamen manuellen Pruefung markiert.
            End If
        End If

        wsCleaned.Cells(i, 5).Value = wsData.Cells(i, 5).Value
        wsCleaned.Cells(i, 6).Value = wsData.Cells(i, 6).Value

    Next i

    For i = 2 To lastRow                               ' Diese zweite Schleife prueft, ob der Execution Price gegen den Limit Price verstoesst (Buy: Execution Price > Limit Price; Sell: Execution
        Dim orderTyp As Variant                        ' Price < Limit Price). Sie kann nicht in die erste Schleife integriert werden, da beide Werte (Limit Price und Execution Price) dort erst
        orderTyp = wsCleaned.Cells(i, 6).Value         ' vollstaendig und korrekt geschrieben sein muessen, bevor der Vergleich stattfinden kann.
        Dim execuPrice As Variant
        execuPrice = wsCleaned.Cells(i, 11).Value
        Dim limitPrice As Variant
        limitPrice = wsCleaned.Cells(i, 7).Value

        If limitPrice <> "" And execuPrice <> "" Then
            If orderTyp = "Limit Order" Then
                If wsCleaned.Cells(i, 5).Value = "Buy" And execuPrice > limitPrice Then
                    LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Buy: Execution Price exceeds Limit Price"
                ElseIf wsCleaned.Cells(i, 5).Value = "Sell" And execuPrice < limitPrice Then
                    LogQualityIssue wsQuality, qualityRow, i, wsData.Cells(i, 1).Value, "Sell: Execution Price below Limit Price"
                End If
            End If
        End If
    Next i

    For Each key In dict.keys
        If InStr(dict(key), ",") > 0 Then
            rowList = Replace(dict(key), ",", ", ")
            qualityRow = qualityRow + 1
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
