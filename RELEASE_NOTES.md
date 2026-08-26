# Plenaro

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
