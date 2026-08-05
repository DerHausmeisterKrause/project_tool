# TaskTool (WPF, .NET 8)

Lokale Aufgaben- und Zeitverwaltungs-App mit SQLite und Outlook Busy-Blocker via COM Interop.

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

Die Ticketsystem-Anbindung nutzt den Znuny/OTRS 6.3.x `GenericTicketConnectorREST` mit `SessionCreate`, `SessionGet`, `TicketSearch` und `TicketGet`. Als Server-URL wird die Basis des Webservice erwartet, zum Beispiel:

`https://SERVER/nph-genericinterface.pl/Webservice/GenericTicketConnectorREST`

Die Anmeldung für `TicketSearch` und `TicketGet` erfolgt standardmäßig direkt pro Operation mit `UserLogin` und `Password`; API-Tokens werden für diesen Connector nicht verwendet. In den Einstellungen muss eine **Znuny Agenten-ID** hinterlegt werden. Diese interne numerische ID wird für `OwnerIDs` und `ResponsibleIDs` in `TicketSearch` verwendet. Die TicketSearch-Route (`POST /Ticket/Search` oder kompatibel `GET /Ticket`) und das TicketGet-Route-Template sind konfigurierbar. `SessionGet` wird im normalen Sync nicht benötigt und ist nur noch Diagnose/Fallback.

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
