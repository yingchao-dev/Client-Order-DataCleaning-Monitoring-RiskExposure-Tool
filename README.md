# Client Order Data Cleaning, Monitoring & Risk Exposure Tool — README

*Portfolio Project — Trading Support & Execution Quality Analysis*

**Autor:** Yingchao Li<br>
**Stand:** 25.09.2026<br>
**Struktur:** 5-Layer-Architektur (Cleaning → Monitoring → Exposure → Reporting, orchestrated by Start_Here)

---

Dieses Projekt bildet einen vereinfachten Kundenauftraege-Ueberwachungsprozess in 5 Schichten ab:

### ① DataCleaning
Validierung, Bereinigung und Anreicherung (Lookup auf Security_Master/Client_Master) der Rohdaten; Ergebnis ist eine bereinigte, zusammengefuehrte Datenbasis (Cleaned_Data); Luecken sind mit Ursache in Worksheet Data_Quality dokumentiert.

### ② Monitoring
Regelbasierte Erkennung auffaelliger Einzelauftraege (offene Pending-Orders, Preisabweichung, Ausfuehrungsverzoegerung, Mengenspitzen).

### ③ AssetExposure
Aggregation von Net-/Gross-Exposure auf Ebene Einzelwert (ISIN) und Instrument-Klasse, inkl. Umrechnung von Bond-Kursen (Prozent vom Nominalwert); zusaetzlich Erkennung von Einzelgeschaeften ueber Grossauftrags-Schwellenwerten (Notional Value).

### ④ Output
Aufbau eines Summary-Sheets (BuildSummary) mit Anzahl auffaelliger Orders je Kategorie (dedupliziert ueber die vier Monitoring-Sheets), Gesamt-Gross-Exposure (aus Layer③ uebernommen, nicht neu berechnet), Top-3-Kunden nach Volumen und Auffaelligkeitsrate je Asset-Klasse (auffaellig/gesamt).

### ⑤ Start_Here
Zentraler Einstiegspunkt (RunFullPipeline), der alle Sub-Prozeduren aus Layer①-④ in der richtigen Reihenfolge aufruft; Zusammenfassung bei Erfolg, Hinweismeldung mit moeglicher Ursache bei Abbruch.

---

## Datensatz-Hintergrund

- **Datenherkunft:** Der im Projekt verwendete Datensatz (Raw_Client_Orders, Security_Master, Client_Master) wurde vollstaendig mittels KI generiert und simuliert. Es handelt sich um keine echten Kunden- oder Handelsdaten; es bestehen daher keine Datenschutz- oder Compliance-Bedenken. Der Datensatz wurde bewusst mit realistischen Datenqualitaetsproblemen versehen (fehlerhafte Formate, ungueltige Referenzen, widerspruechliche Zeitstempel etc.), um die Cleaning- und Validierungslogik gezielt zu testen.
- Raw_Client_Orders wird als untrusted Quelle behandelt und nicht manuell vorab bereinigt; alle Korrekturen laufen ueber die Cleaning-Logik in Layer①.
- Security_Master und Client_Master werden als vertrauenswuerdige Stammdaten behandelt und nicht bereinigt; das Cleaning beschraenkt sich ausschliesslich auf Raw_Client_Orders. Diese Annahme ist in der Praxis plausibel, da Stammdaten typischerweise durch einen separaten, staerker kontrollierten Prozess gepflegt werden.
- Da diese Datei ueber das Internet heruntergeladen wurde, blockiert Windows moeglicherweise standardmaessig die Makroausfuehrung. Nach dem Download: Rechtsklick auf die Datei → Eigenschaften → im Reiter "Allgemein" unten die Option "Blockierung aufheben" aktivieren → Uebernehmen und OK. Anschliessend beim Oeffnen in Excel "Inhalt aktivieren" waehlen.
- Gelb markierte Zeilen in Cleaned_Data kennzeichnen Auftraege, deren Einzelposition-Exposure den definierten Schwellenwert ueberschreitet. Fuer eine Erklaerung den Mauszeiger auf eine beliebige Stelle der Kopfzeile bewegen.

## Design-Entscheidungen

- **VBA statt Power Query fuer die Cleaning-/Monitoring-Logik:** die Validierungsregeln pruefen mehrere Felder gegeneinander (z.B. Status vs. Execution Price/Time, Order Type vs. Limit Price) und schreiben strukturierte Fehlermeldungen in ein separates Audit-Sheet - das laesst sich in VBA klarer und wartbarer abbilden als in Power Query/M.
- **Application.ScreenUpdating und Application.Calculation** werden waehrend der Pipeline-Ausfuehrung deaktiviert: bei zeilenweiser Verarbeitung (aktuell 2.000 Zeilen) reduziert dies die Laufzeit in gewissem Umfang, da Bildschirm-Neuzeichnung und automatische Neuberechnung bei jeder einzelnen Zellaenderung entfallen. Bei moderat groesseren Datenmengen (z.B. im Bereich mehrerer Zehntausend Zeilen) waere als naechster Schritt eine array-basierte Verarbeitung denkbar (Batch-Read/Write statt zeilenweiser Cell-Zugriffe). Bei einer weiteren Groessenordnung (mehrere Hunderttausend Zeilen) waere VBA/Excel jedoch kein geeignetes Werkzeug mehr - eine Datenverarbeitung in diesem Umfang gehoert eher in eine Python-basierte Loesung (z.B. mit pandas fuer statistische Auswertung), die fuer grosse Datenmengen und komplexere Berechnungen besser geeignet ist als VBA/Excel. Dieses Projekt ist als kleines Tool fuer den Front-Office-Einsatz konzipiert und verarbeitet entsprechend begrenzte Datenmengen; die aktuelle zeilenweise Verarbeitung mit leichter Performance-Optimierung kann die praktischen Anforderungen vorerst abdecken.

---

Abschliessend folgt hier eine "Schnelluebersicht", die dabei hilft, sich schnell einen Ueberblick ueber den Inhalt jedes neu erstellten Sheets sowie die zugehoerige Code-Position zu verschaffen.

| Sheet | Inhalt | Quelle (Layer / Sub) |
|---|---|---|
| Cleaned_Data | Bereinigte Auftragsdaten (Luecken siehe Data_Quality) | Layer① / DataCleaning |
| Data_Quality | Liste der erkannten Datenqualitaetsprobleme | Layer① / DataCleaning |
| Pending_Monitoring | Lange unbearbeitete Pending-Auftraege | Layer② / MonitorPendingOrders |
| Price_Deviation_Monitoring | Auftraege mit Preisabweichung vom Wertpapier-Durchschnittspreis (ohne sich selbst) | Layer② / MonitorPriceDeviation |
| Execution_Delay_Monitoring | Auftraege mit ueberhoehter Ausfuehrungsverzoegerung | Layer② / MonitorExecutionDelay |
| Order_Quantity_Spike_Monitoring | Auftraege mit ungewoehnlichem Mengenanstieg | Layer② / MonitorOrderQuantitySpike |
| Large_Trade_Alert | Einzelgeschaefte ueber dem Exposure-Schwellenwert | Layer③ / MonitorLargeTrade |
| InstrumentTyp_Exposure | Exposure aggregiert nach Asset-Klasse | Layer③ / MonitorAssetExposure-InstrumentType |
| Security_Exposure | Exposure aggregiert nach Einzelwert (ISIN) | Layer③ / MonitorAssetExposure-BySecurity |
| Summary | Taegliche Zusammenfassung und Kennzahlen | Layer④ / BuildSummary |

*Raw_Client_Orders, Security_Master und Client_Master sind die urspruenglichen Eingabedaten und werden von RunFullPipeline nicht veraendert.*
