# Plenaro V2.2.4

## Bugfixes

- Buildfehler in der Verarbeitung wartender Znuny-Tickets behoben.
- Falsch benannter Parameter beim Speichern verarbeiteter Wartetermine korrigiert.

# Plenaro V2.2.3

## Wartende Znuny-Tickets

- Kritischen Fehler bei der Berechnung von Znuny-Wartezeiten behoben.
- Znuny `UntilTime` wird jetzt als vorzeichenbehaftete verbleibende Zeit in Sekunden interpretiert und nicht mehr als Unix-Zeitstempel aus dem Jahr 1970.
- Absolute `PendingTime`-Werte und relative `UntilTime`-Werte werden getrennt verarbeitet; ein gültiger absoluter Wert hat Vorrang.
- Pending-Benachrichtigungen starten erst nach einem erfolgreichen initialen Znuny-Abgleich.
- Bereits fällige Wartetermine werden beim Start still als Baseline übernommen; alte lokale Daten erzeugen keine rückwirkenden Benachrichtigungen.
- Nur in der laufenden, erfolgreich synchronisierten Sitzung fällig werdende und remote bestätigte `pending reminder`-Tickets können einmalig benachrichtigen.
- Ein Circuit Breaker deaktiviert Pending-Benachrichtigungen für die Sitzung bei Massenauslösungen oder wiederholten Schedulerfehlern, ohne den normalen Ticket-Sync zu stoppen.
- Aktuelle Aufgaben werden weiterhin sofort aus SQLite geladen; der Znuny-Initial-Sync beginnt danach im Hintergrund.
- Geschlossene, nicht mehr zugewiesene oder zeitlich nicht eindeutig auflösbare Tickets erzeugen keine Pending-Benachrichtigung.
