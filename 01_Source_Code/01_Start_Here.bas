Attribute VB_Name = "Start_Here"
Option Explicit
' ============================================================
' Layer 5: Start_Here / RunFullPipeline
' Zweck: Zentraler Einstiegspunkt fuer den gesamten Ablauf. Ruft
'        alle Sub-Prozeduren aus Layer1-4 in der richtigen
'        Reihenfolge auf (Cleaning ¡ú Monitoring ¡ú Exposure Tracking
'        ¡ú Reporting) und zeigt am Ende eine Zusammenfassung der
'        erstellten Sheets an.
' Ausfuehrung: Cursor an eine beliebige Stelle innerhalb dieser
'        Prozedur setzen, dann F5 oder den gruenen "Start"-Pfeil
'        in der VBA-Toolbar klicken.
' Input: Raw_Client_Orders, Security_Master, Client_Master (muessen
'        vollstaendig und unveraendert vorliegen)
' Output: Alle in Layer1-4 definierten Sheets (insgesamt 10 neue,
'        zusaetzlich zu den 4 urspruenglichen)
' Fehlerbehandlung: Bei einem Abbruch waehrend der Ausfuehrung wird
'        eine Hinweismeldung mit moeglicher Ursache angezeigt statt
'        eines technischen VBA-Fehlerdialogs.
' ============================================================
Sub RunFullPipeline()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Call DataCleaning
    Call MonitorPendingOrders
    Call MonitorPriceDeviation
    Call MonitorExecutionDelay
    Call MonitorOrderQuantitySpike
    Call MonitorLargeTrade
    Call MonitorAssetExposureInstrumentType
    Call MonitorAssetExposureBySecurity
    Call BuildSummary
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    
    MsgBox "Programm erfolgreich ausgefuehrt. Neben den urspruenglichen 4 Sheets " & _
           "liegen jetzt insgesamt 14 Arbeitsblaetter vor.", vbInformation
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True             ' bei Abbruch ebenfalls zuruecksetzen, sonst bleibt Excel dauerhaft im manuellen Modus
    Application.Calculation = xlCalculationManual  ' bei Abbruch ebenfalls zuruecksetzen, sonst bleibt der Bildschirm dauerhaft eingefroren
    MsgBox "Der Ablauf wurde unterbrochen. Bitte pruefen Sie, ob alle " & _
           "erforderlichen Sheets (Raw_Client_Orders, Security_Master, " & _
           "Client_Master) vorhanden und unveraendert sind, und starten " & _
           "Sie RunFullPipeline erneut." & vbNewLine & vbNewLine & _
           "Technische Details: " & Err.Description
End Sub



