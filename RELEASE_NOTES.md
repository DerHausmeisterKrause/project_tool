# Plenaro

## Znuny / OTRS

- Ticket-Synchronisierung auf eine lokale Cache-/Snapshot-Architektur umgestellt
- Beim Programmstart wird genau ein kontrollierter Znuny-Abgleich durchgeführt
- Ticketdaten werden anschließend lokal gespeichert und von der Oberfläche aus dem lokalen Cache verwendet
- Automatische Aktualisierungen erfolgen ausschließlich nach dem in den Einstellungen festgelegten Synchronisationsintervall
- Separate automatische Candidate- und Hintergrund-Refreshes entfernt
- Ansichtswechsel und Einstellungsänderungen lösen keine zusätzlichen Znuny-Anfragen mehr aus
- „Aktuelle Aufgaben“ und „Neue Aufgaben“ verwenden zwischen Synchronisierungen ausschließlich lokal gespeicherte Daten
- Candidate-/Neue-Aufgaben-Daten werden lokal zwischengespeichert
- Historische Tickets werden nicht mehr bei jedem Synchronisationslauf erneut vollständig vom Server abgefragt
- Hintergrund-Synchronisierung lädt nur noch tatsächlich benötigte Ticketdaten
- Schwere Ticketdetails und Artikel werden nicht unnötig bei jedem Sync erneut geladen
- Automatische Fehler führen nicht mehr zu schnellen Wiederholungsanfragen
- Explizite Benutzeraktionen bleiben weiterhin möglich, lösen jedoch keinen unnötigen vollständigen Hintergrund-Sync aus
- Mindestintervall für automatische Ticket-Synchronisierung auf fünf Minuten begrenzt
- Zusätzliche Sicherheitsgrenze verhindert unerwartet große automatische Request-Mengen
- Neue Diagnose-Logs zeigen die tatsächliche Anzahl von Znuny-Anfragen pro Synchronisationslauf

## Favoriten & eingebettete Webseiten

- Fehler beim Wechsel zwischen mehreren eingebetteten Webseiten behoben
- WebView2-Fehler bei unterschiedlichen Browser-Umgebungen behoben
- Beim Wechsel zwischen Favoriten wird zuverlässig die richtige Webseite angezeigt
- Race Conditions bei schnellen Wechseln zwischen mehreren Webseiten beseitigt
- Bereits geöffnete Webseiten bleiben im Hintergrund aktiv
- Browserzustand, aktuelle Unterseite, Scrollposition, JavaScript-State und Login-Sessions bleiben beim Wechsel erhalten
- Webseiten werden weiterhin erst beim ersten Öffnen geladen
- Jeder Favorit verwendet ein eigenes isoliertes WebView2-Profil
- CORS-/Browser-Websecurity kann gezielt pro Favorit deaktiviert werden
- Optionale Unterstützung für Mixed Content ergänzt
- CORS- und Mixed-Content-Ausnahmen gelten ausschließlich für den jeweiligen Favoriten
- Netzwerk- und API-Fehler eingebetteter Webseiten können genauer diagnostiziert werden
- AutoLogin bleibt auf den konfigurierten HTTPS-Host beschränkt
- Favoriten können frei sortiert werden
- Neue Funktionen zum Verschieben eines Favoriten nach oben oder unten
- Benutzerdefinierte Favoriten-Reihenfolge wird dauerhaft gespeichert
- Sortierung wird direkt in der Navigation übernommen
- Umbenennen oder Sortieren eines Favoriten lädt dessen Webseite nicht unnötig neu

## Segmentplanung

- Neue visuelle Tages-Verfügbarkeitsanzeige in der Segmentplanung
- Der Zeitraum von 06:00 bis 18:00 wird als kompakter Zeitbalken dargestellt
- Aufteilung in 48 einzelne 15-Minuten-Segmente
- Freie Zeiträume werden grün dargestellt
- Durch Outlook-Termine belegte Zeiträume werden rot dargestellt
- Jede volle Stunde wird direkt an der Zeitleiste beschriftet
- Die Anzeige folgt automatisch dem ausgewählten Planungsdatum
- Outlook-Kalenderdaten werden effizient über den bestehenden Cache verwendet
- Fehlende Kalenderzeiträume können im Hintergrund nachgeladen werden
- Änderungen nach einer Outlook-Synchronisierung aktualisieren die Anzeige automatisch
- Abgesagte und als „Frei“ markierte Termine blockieren keine Zeit
- Ganztagstermine blockieren nicht fälschlicherweise den gesamten Arbeitstag
- Überschneidungen werden korrekt auf die betroffenen 15-Minuten-Segmente abgebildet
- Nicht verfügbare Kalenderdaten werden neutral dargestellt
- Die Anzeige dient ausschließlich als Planungshilfe

## Einstellungen

- Einstellungsbereich grundlegend überarbeitet und übersichtlicher strukturiert
- Neues horizontales Menüband für Allgemein, Aufgaben & Zeiten, Outlook, Homeoffice, Ticketsystem, Wiki, Favoriten und Updates
- Es wird nur noch der aktuell ausgewählte Einstellungsbereich angezeigt
- Umfangreiche Bereiche wurden in thematisch passende Cards gegliedert
- Bestehende Zugangsdaten und Konfigurationen bleiben erhalten
- Darstellung für verschiedene Fenstergrößen und DPI-Skalierungen verbessert
- Settings-Oberfläche in kleinere wartbare Teilansichten aufgeteilt

## Updates

- Automatische Updateprüfung beim Start bleibt erhalten
- Updates werden beim Programmstart nicht mehr ungefragt installiert
- Neuer Hinweisdialog informiert über verfügbare Plenaro-Versionen
- Benutzer kann zwischen „Jetzt installieren“ und „Später“ wählen
- Stille automatische Updateinstallation entfernt
- Derselbe Update-Dialog wird innerhalb einer Sitzung nicht mehrfach angezeigt
- Stable- und Test-Updatekanäle sowie Download-, Prüfsummen- und Rollback-Funktionen bleiben erhalten

## Stabilität & Bedienung

- WebView2-Lifecycle und asynchrone Navigation robuster gestaltet
- Browser-Sessions verschiedener Favoriten stärker voneinander isoliert
- Unnötige Browser-Neuinitialisierungen reduziert
- Kalender- und Update-Hintergrundprozesse benutzerfreundlicher gestaltet
