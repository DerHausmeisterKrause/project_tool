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
- Für maximale Build-Stabilität ist aktuell **kein compile-time ApplicationIcon** im `.csproj` gesetzt.
- Optional kann `Assets/Plenaro.ico` neben der EXE liegen (`<publish>/Assets/Plenaro.ico`) — das Fenster lädt dieses Icon zur Laufzeit (siehe `MainWindow.xaml.cs`).
- **Ohne** Icon-Datei: Build funktioniert weiterhin, es wird das Standard-Icon verwendet.

Hinweis zu `CS7065` ("Symbol-Stream weist nicht das erwartete Format auf"):
- Dieser Fehler entsteht bei ungültigen `.ico` Dateien, wenn sie als `ApplicationIcon` kompiliert werden.
- Deshalb wurde die compile-time Icon-Einbindung entfernt.


### Build-Fehler CS7065 (Win32-Ressourcen)
Falls weiterhin `CS7065` auftritt:
1. `bin/` und `obj/` im Projektordner löschen.
2. `dotnet restore` und `dotnet build -c Release` erneut ausführen.

Hinweis: Das Projekt deaktiviert die Win32-Manifest-Erzeugung (`NoWin32Manifest=true`), damit fehlerhafte lokale Win32/Icon-Resource-Streams den Build nicht mehr blockieren.
