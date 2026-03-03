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
Wenn eure Umgebung keine Binärdateien im Repo erlaubt, legt das Icon lokal ab:

1. Erstelle den Ordner `Assets/` im Projektroot (falls nicht vorhanden).
2. Lege die Datei `Assets/Plenaro.ico` ab (empfohlen: 16/32/48/256 px).
3. Build/Publish wie gewohnt ausführen.

Verhalten:
- **Mit** `Assets/Plenaro.ico` im Projektroot vor dem Build: wird in die App eingebettet und als Application-Icon (EXE/Taskleiste) sowie Fenster-Icon genutzt.
- Alternativ kann `Assets/Plenaro.ico` auch neben der EXE liegen (`<publish>/Assets/Plenaro.ico`) — das Fenster lädt dieses Icon zur Laufzeit als Fallback.
- **Ohne** Icon-Datei: Build funktioniert weiterhin, es wird das Standard-Icon verwendet.
