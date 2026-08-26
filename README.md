# TaskTool (WPF, .NET 8)

Lokale Aufgaben- und Zeitverwaltungs-App mit SQLite und Outlook Busy-Blocker via COM Interop.

## Release Notes

### Updates

- Automatische Updateprüfungen beim Start bleiben erhalten, installieren Updates aber niemals mehr still
- Ein neuer Plenaro-Dialog zeigt installierte und verfügbare Version sowie einen klaren Neustart-Hinweis
- Benutzer können zwischen „Jetzt installieren“ und „Später“ wählen
- Zurückgestellte Updates bleiben im Update-Banner und unter Einstellungen → Updates verfügbar
- Derselbe Dialog erscheint pro App-Sitzung nur einmal; stündliche Prüfungen aktualisieren ausschließlich den Banner
- Alte `AutoInstallUpdatesOnStartup`-Werte werden sicher deaktiviert
- Stable-/Testkanäle, SemVer-Auswahl, SHA-Prüfung, Rollback und Update-Replacement bleiben unverändert

### Wartende Znuny-Tickets

- Wartende Znuny-/OTRS-Tickets mit `pending reminder` oder `pending auto` können optional aus „Heute“ und „Aktuelle Aufgaben“ ausgeblendet werden
- Neue Ansicht „Wartend“ und neue Einstellungen unter **Ticketsystem → Wartende Tickets**
- Znuny-Wartezeiten werden lokal gespeichert und nach einem Neustart unmittelbar nachgeprüft
- Ein lokaler Einzeltimer plant jeweils den nächsten Wartetermin und führt danach ein gezieltes TicketGet aus
- Verlängerte Wartezeiten werden neu eingeplant; geschlossene oder anders zugewiesene Tickets erzeugen keine Benachrichtigung
- Wake-Benachrichtigungen werden pro Wartetermin persistent dedupliziert und bei vielen gleichzeitig fälligen Tickets zusammengefasst

### Einstellungen

- Einstellungsbereich mit einem dauerhaft sichtbaren horizontalen Menüband neu strukturiert
- Kategorien für Allgemein, Aufgaben & Zeiten, Outlook, Homeoffice, Ticketsystem, Wiki, Favoriten und Updates eingeführt
- Nur die aktive Kategorie wird geladen und in thematisch gruppierten Cards dargestellt
- Webseiten-Shortcuts werden in den Einstellungen konsistent als Favoriten bezeichnet
- Update-Hinweise öffnen direkt die Update-Kategorie
- Bestehende Bindings, verschlüsselte Zugangsdaten, Wiki-Quellen und Favoriten-Sortierung bleiben erhalten
- Responsive Navigation, begrenzte Inhaltsbreite und unabhängiges Scrollen der aktiven Kategorie verbessern die Nutzung bei unterschiedlichen Fenstergrößen und DPI-Skalierungen

### Segmentplanung

- Neue visuelle Tages-Verfügbarkeitsanzeige von 06:00 bis 18:00 mit 48 proportionalen 15-Minuten-Segmenten
- Freie Zeiträume werden grün, durch Outlook-Termine belegte Zeiträume rot und nicht verfügbare Kalenderdaten neutral dargestellt
- Die Anzeige folgt dem ausgewählten Segmentdatum und aktualisiert sich nach Outlook-Synchronisationen automatisch
- Abgesagte, als frei markierte und ganztägige Termine blockieren die Anzeige nicht

**Plenaro 2.1.0**

- Neue Ausschlusswörter für „Neue Aufgaben“
- Tickets können trotz passendem Suchwort gezielt über unerwünschte Begriffe ausgefiltert werden
- Ausschlussfilter berücksichtigt Titel, Nachrichtenbetreff und Nachrichteninhalt
- „Mir zuweisen“ deutlich beschleunigt
- Nach Self-Assign wird nur noch das betroffene Ticket gezielt aktualisiert
- Unnötiger vollständiger Ticket- und Candidate-Sync nach jeder Zuweisung entfällt
- Zugewiesene Candidates verschwinden unmittelbar aus „Neue Aufgaben“
- Zugewiesene Tickets erscheinen ohne langen Vollsync unter „Aktuelle Aufgaben“
- Bestehende Assignment-Notifications und deren Self-Assign-Unterdrückung bleiben erhalten

## NuGet Pakete
- `Microsoft.Data.Sqlite` (SQLite-Datei DB)
- `Microsoft.Office.Interop.Outlook` (COM Interop für Outlook-Termine)

## Build
```bash
dotnet restore
dotnet build -c Release
```

## Publish (Portable Single EXE, win-x64, self-contained)
```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true /p:PublishTrimmed=false
```

Output liegt unter:
`bin/Release/net8.0-windows/win-x64/publish/`


## Windows-EXE erstellen

Eine veröffentlichungsfertige Windows-x64-EXE kann direkt über das Release-Skript im Projekt-Root erzeugt werden:

- Rechtsklick auf `build-release.ps1` → **Mit PowerShell ausführen**
- oder im Terminal:

```powershell
powershell -ExecutionPolicy Bypass -File .\build-release.ps1
```

Das Skript führt `dotnet clean`, `dotnet restore` und `dotnet publish` für `TaskTool.Wpf.csproj` aus. Die fertige EXE liegt anschließend unter:

`artifacts\publish\win-x64`


## Znuny / OTRS Ticketsystem

Die Ticketsystem-Anbindung nutzt den Znuny/OTRS 6.3.x `GenericTicketConnectorREST` im normalen Sync mit `SessionCreate`, `TicketSearch` und `TicketGet`; `SessionGet` ist nur Diagnose/Fallback. Als Server-URL wird die Basis des Webservice erwartet, zum Beispiel:

`https://SERVER/nph-genericinterface.pl/Webservice/GenericTicketConnectorREST`

Die Anmeldung erfolgt standardmäßig über `POST /Session`; anschließend werden `TicketSearch` und `TicketGet` mit `SessionID` ausgeführt, damit Benutzername und Passwort nicht in GET-URLs landen. API-Tokens werden für diesen Connector nicht verwendet. In den Einstellungen muss eine **Znuny Agenten-ID** hinterlegt werden. Diese interne numerische ID wird für `OwnerIDs` und `ResponsibleIDs` in `TicketSearch` verwendet. Für ältere OTRS/Znuny-6.x-GenericTicketConnectorREST-Konfigurationen ist `GET /Ticket` die Standardroute für TicketSearch; `POST /Ticket/Search` bleibt als optionale neuere Variante konfigurierbar. `SessionGet` wird im normalen Sync nicht benötigt und ist nur noch Diagnose/Fallback.

## Start
1. Aktuelle `TaskTool.zip` aus dem Ordner "deploy" entpacken.
2. `TaskTool.exe` starten.
3. Beim ersten Start werden automatisch erzeugt (neben der EXE):
   - `TaskTool.db`
   - `settings.json`
   - `logs.txt`
4. Outlook-Integration kann in **Einstellungen** deaktiviert werden.

## Hinweise
- Alles bleibt lokal, keine Cloud, kein Webserver.
- Outlook-Reminder werden immer deaktiviert (`ReminderSet = false`).
- Bei Outlook/COM Fehlern läuft die App weiter; Fehlertext erscheint in der Heute-Ansicht und im Log.

## License
MIT
