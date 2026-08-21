# TaskTool (WPF, .NET 8)

Lokale Aufgaben- und Zeitverwaltungs-App mit SQLite und Outlook Busy-Blocker via COM Interop.

## Release Notes

**Plenaro 2.1.0**

- Sortierung von Znuny-Tickets nach echtem Ticket-Erstellungsdatum korrigiert
- „Zuletzt bearbeitet“ berücksichtigt jetzt tatsächliche Ticket- und Plenaro-Aktivitäten
- Segment- und Terminplanung zählt jetzt als Bearbeitung
- Ticketantworten, Zeitbearbeitung und erfolgreiche Zeitbuchungen fließen in die Aktivitätssortierung ein
- Automatische Synchronisierungen verfälschen die Reihenfolge nicht mehr
- Angepinnte Aufgaben bleiben unabhängig von der Sortierung weiterhin ganz oben

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
