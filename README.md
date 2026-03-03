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

## Start
1. `TaskTool.exe` starten.
2. Beim ersten Start werden automatisch erzeugt (neben der EXE):
   - `TaskTool.db`
   - `settings.json`
   - `logs.txt`
3. Outlook-Integration kann in **Einstellungen** deaktiviert werden.

## Hinweise
- Alles bleibt lokal, keine Cloud, kein Webserver.
- Outlook-Reminder werden immer deaktiviert (`ReminderSet = false`).
- Bei Outlook/COM Fehlern läuft die App weiter; Fehlertext erscheint in der Heute-Ansicht und im Log.

## Optional: Eigenes App/Fenster-Icon (manuell)
Lege (lokal) die Datei `Assets/Plenaro.ico` im Projektroot ab.

Was dann automatisch passiert:
- Die Datei wird in Output/Publish mitkopiert (`CopyToOutputDirectory` + `CopyToPublishDirectory`).
- Das Fenster-Icon wird zur Laufzeit geladen (`MainWindow.xaml.cs`) über:
  1. eingebettete Resource (`pack://application:,,,/Assets/Plenaro.ico`)
  2. Fallback auf `<publish>/Assets/Plenaro.ico`.

Wichtig zu `CS7065` ("Symbol-Stream weist nicht das erwartete Format auf"):
- Dieser Fehler kommt von einer ungültigen `.ico` Datei (z. B. PNG nur umbenannt in `.ico`).
- Deshalb ist das compile-time EXE-Icon standardmäßig deaktiviert, damit Builds stabil laufen.

Optional EXE/Taskleisten-Icon aktivieren (nur mit **valider** `.ico`):
```bash
dotnet build -c Release -p:EnableCompileTimeAppIcon=true
```

Empfehlung für die Icon-Datei: echtes Multi-Size ICO (mind. 16, 32, 48, 256 px).
