# Plenaro V2.2.5

## Znuny / OTRS – Stabilität

- Automatische Znuny-/OTRS-Anfragen umfassend gegen Serverüberlastung abgesichert.
- Dedizierten Netzwerk-Wake-Timer für wartende Tickets entfernt.
- Wartende Tickets werden jetzt ausschließlich über den kontrollierten regulären Znuny-Sync aktualisiert.
- Sekundenbasierte Wiederholungen bei wartenden Tickets vollständig entfernt.
- Historische Znuny-Tickets werden nicht mehr bei jedem Synchronisationslauf erneut vollständig vom Server geladen.
- Hintergrund-Synchronisierung verwendet nach Möglichkeit schlanke Ticket-Metadaten statt sämtlicher Artikel und Dynamic Fields.
- Candidate-Synchronisierung vom normalen Aufgaben-Sync entkoppelt und deutlich serverfreundlicher gestaltet.
- Parallele schwere Ticket-Abfragen stark begrenzt.
- Globales Rate-Limit für automatische Znuny-/OTRS-Anfragen eingeführt.
- Globaler Circuit Breaker pausiert automatische Synchronisierung bei ungewöhnlich hoher Anfragezahl.
- Automatische Fehlerwiederholungen verwenden ausschließlich kontrollierten Backoff.
- Mindestintervall für automatische vollständige Ticket-Synchronisierung erhöht.
- Diagnose-Logging zeigt die Anzahl automatischer API-Anfragen ohne sensible Daten.
- Lokale Aufgaben bleiben während aller Hintergrund-Synchronisierungen sofort verfügbar.

# Plenaro V2.2.4

## Bugfixes

- Buildfehler in der Verarbeitung wartender Znuny-Tickets behoben.
- Falsch benannter Parameter beim Speichern verarbeiteter Wartetermine korrigiert.

## Aufgaben

- Fehler behoben, durch den „Aktuelle Aufgaben“ nach dem Programmstart zunächst leer erscheinen konnten.
- Lokal gespeicherte Aufgaben werden jetzt unabhängig vom noch ausstehenden Znuny-Sync sofort angezeigt.
- Veraltete lokale Zuweisungsinformationen blenden Znuny-Tickets während des Startvorgangs nicht mehr vorzeitig aus.
- Owner- und Responsible-Zuweisungen werden erst nach dem ersten erfolgreichen Znuny-Abgleich als verbindlicher Sichtbarkeitsfilter verwendet.
- Nach Abschluss der Znuny-Synchronisierung werden die Aufgabenlisten automatisch mit dem aktuellen Remote-Zustand abgeglichen.
- Lokale Aufgaben bleiben auch bei langsamem oder nicht erreichbarem Znuny sofort sichtbar.
- Rückkehr zur Heute-Ansicht aktualisiert die Aufgabenliste zusätzlich zuverlässig.

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
