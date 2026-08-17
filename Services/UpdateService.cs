using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class UpdateService : IDisposable
{
    public const string LatestReleaseEndpoint = "https://api.github.com/repos/DerHausmeisterKrause/project_tool/releases/latest";
    private const string AssetName = "TaskTool.exe";
    private readonly LoggerService _logger;
    private readonly SettingsService _settings;
    private readonly AppVersionService _version;
    private readonly HttpClient _client = new() { Timeout = TimeSpan.FromSeconds(30) };
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly CancellationTokenSource _shutdown = new();

    public UpdateService(LoggerService logger, SettingsService settings, AppVersionService version)
    {
        _logger = logger;
        _settings = settings;
        _version = version;
        var userAgentVersion = version.TryGetInstalledVersion(out var installedVersion) ? installedVersion.ToString() : "unknown";
        _client.DefaultRequestHeaders.UserAgent.ParseAdd($"Plenaro-Updater/{userAgentVersion}");
        _client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
    }

    public async Task<UpdateCheckResult> CheckForUpdatesAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            if (!_version.TryGetInstalledVersion(out var installedVersion))
                throw new InvalidOperationException("Installierte Versionsinformation ist ungültig.");
            using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _shutdown.Token);
            using var response = await _client.GetAsync(LatestReleaseEndpoint, linked.Token);
            if (response.StatusCode == HttpStatusCode.Forbidden) throw new InvalidOperationException("GitHub Rate Limit erreicht (HTTP 403).");
            if (response.StatusCode == HttpStatusCode.NotFound) throw new InvalidOperationException("GitHub Release wurde nicht gefunden (HTTP 404).");
            response.EnsureSuccessStatusCode();
            await using var stream = await response.Content.ReadAsStreamAsync(linked.Token);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: linked.Token);
            var root = doc.RootElement;
            var tag = Read(root, "tag_name");
            var remote = AppVersionService.ParseVersion(tag);
            var assets = root.GetProperty("assets").EnumerateArray();
            var assetElement = assets.FirstOrDefault(a => string.Equals(Read(a, "name"), AssetName, StringComparison.OrdinalIgnoreCase));
            if (assetElement.ValueKind == JsonValueKind.Undefined) throw new InvalidOperationException("Release-Asset TaskTool.exe fehlt.");
            var asset = new GitHubReleaseAsset(Read(assetElement, "name"), Read(assetElement, "browser_download_url"), assetElement.GetProperty("size").GetInt64(), Read(assetElement, "digest"));
            var info = new UpdateInfo(remote, tag, Read(root, "name"), Read(root, "body"), Read(root, "html_url"), root.TryGetProperty("published_at", out var p) && p.TryGetDateTimeOffset(out var published) ? published : null, asset);
            var available = remote > installedVersion;
            _logger.Info($"[UpdateCheck] installedVersion={installedVersion} remoteVersion={remote} updateAvailable={available.ToString().ToLowerInvariant()}");
            return new UpdateCheckResult(available, available ? info : null, available ? $"Plenaro {remote} ist verfügbar." : "Kein Update verfügbar.");
        }
        catch (Exception ex) { _logger.Error($"[UpdateCheck] failed={ex.Message}"); throw; }
        finally { _gate.Release(); }
    }

    public async Task<string> DownloadUpdateAsync(UpdateInfo info, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(info.Asset.Digest) || !info.Asset.Digest.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Das Release enthält keinen verwendbaren SHA256-Digest.");
        var folder = Path.Combine(Path.GetTempPath(), "PlenaroUpdate", info.Version.ToString());
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "TaskTool.exe.new");
        _logger.Info($"[UpdateDownload] version={info.Version} asset='{info.Asset.Name}' size={info.Asset.Size}");
        using var response = await _client.GetAsync(info.Asset.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        response.EnsureSuccessStatusCode();
        await using (var input = await response.Content.ReadAsStreamAsync(cancellationToken))
        await using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
        {
            var buffer = new byte[81920]; long total = 0; int read;
            while ((read = await input.ReadAsync(buffer, cancellationToken)) > 0)
            {
                await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken); total += read;
                if (info.Asset.Size > 0) progress?.Report((int)Math.Clamp(total * 100 / info.Asset.Size, 0, 100));
            }
        }
        await using var verificationStream = File.OpenRead(path);
        var actual = Convert.ToHexString(await SHA256.HashDataAsync(verificationStream, cancellationToken));
        var expected = info.Asset.Digest[7..].Trim();
        var matches = string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        _logger.Info($"[UpdateVerify] sha256Match={matches.ToString().ToLowerInvariant()}");
        if (!matches) { File.Delete(path); throw new InvalidOperationException("SHA256-Prüfung des Updates fehlgeschlagen."); }
        return path;
    }

    public bool InstallUpdate(string sourcePath, UpdateInfo info, out string error)
    {
        error = string.Empty;
        try
        {
            if (!_version.TryGetInstalledVersion(out var currentVersion))
                throw new InvalidOperationException("Installierte Versionsinformation ist ungültig.");
            var target = Environment.ProcessPath ?? throw new InvalidOperationException("Pfad der laufenden EXE ist unbekannt.");
            var backup = target + ".old";
            var cmd = Path.Combine(Path.GetDirectoryName(sourcePath)!, "update.cmd");
            File.WriteAllText(cmd, BuildCommandScript(Environment.ProcessId, sourcePath, target, backup, info.Version.ToString()), new UTF8Encoding(false));
            var elevate = !CanWriteDirectory(Path.GetDirectoryName(target)!);
            _logger.Info($"[UpdateInstall] currentVersion={currentVersion} targetVersion={info.Version} source='{sourcePath}' target='{target}' pid={Environment.ProcessId}");
            var psi = new ProcessStartInfo(cmd) { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden };
            if (elevate) psi.Verb = "runas";
            Process.Start(psi);
            _logger.Info("[UpdateInstall] cmdCreated=true");
            return true;
        }
        catch (Exception ex) { error = ex.Message; _logger.Error($"[UpdateInstall] failed={ex.Message}"); return false; }
    }

    public bool CompletePostUpdate(string targetVersion)
    {
        _logger.Info($"[UpdateLaunch] postUpdateVersion={targetVersion}");
        if (!_version.ApplyPostUpdateVersion(targetVersion)) return false;
        var target = Environment.ProcessPath; if (string.IsNullOrWhiteSpace(target)) return false;
        var success = true;
        try
        {
            var backup = target + ".old"; if (File.Exists(backup)) File.Delete(backup);
            var currentTemp = Path.Combine(Path.GetTempPath(), "PlenaroUpdate", AppVersionService.ParseVersion(targetVersion).ToString());
            if (Directory.Exists(currentTemp)) Directory.Delete(currentTemp, true);
            CleanupOldTempFolders();
        }
        catch (Exception ex) { success = false; _logger.Error($"[PostUpdate] cleanup={ex.Message}"); }
        _logger.Info($"[PostUpdate] installedVersion={_settings.Current.InstalledVersion} cleanupSuccess={success.ToString().ToLowerInvariant()}");
        return true;
    }

    public void CleanupOldTempFolders()
    {
        var root = Path.Combine(Path.GetTempPath(), "PlenaroUpdate"); if (!Directory.Exists(root)) return;
        foreach (var dir in Directory.GetDirectories(root)) try { if (Directory.GetCreationTimeUtc(dir) < DateTime.UtcNow.AddDays(-7)) Directory.Delete(dir, true); } catch { }
    }

    private static string BuildCommandScript(int pid, string source, string target, string backup, string targetVersion)
    {
        static string E(string value) => value.Replace("%", "%%", StringComparison.Ordinal);
        return string.Join("\r\n", new[]
        {
            "@echo off", "setlocal DisableDelayedExpansion", $"set \"PID={pid}\"",
            $"set \"SOURCE={E(source)}\"", $"set \"TARGET={E(target)}\"", $"set \"BACKUP={E(backup)}\"", $"set \"TARGETVERSION={E(targetVersion)}\"",
            "for /L %%I in (1,1,60) do (", "  tasklist /FI \"PID eq %PID%\" /NH | findstr /R /C:\"[ ]%PID%[ ]\" >nul || goto stopped",
            "  timeout /t 1 /nobreak >nul", ")", "exit /b 1", ":stopped", "if not exist \"%SOURCE%\" goto abort",
            "if exist \"%BACKUP%\" del /F /Q \"%BACKUP%\"", "move /Y \"%TARGET%\" \"%BACKUP%\" || goto abort",
            "move /Y \"%SOURCE%\" \"%TARGET%\" || goto rollback", "start \"\" \"%TARGET%\" --post-update-version \"%TARGETVERSION%\"", "goto done",
            ":rollback", "if exist \"%TARGET%\" del /F /Q \"%TARGET%\"", "move /Y \"%BACKUP%\" \"%TARGET%\"",
            "start \"\" \"%TARGET%\"", "exit /b 4", ":abort", "start \"\" \"%TARGET%\"", "exit /b 5",
            ":done", "endlocal", "del \"%~f0\"", string.Empty
        });
    }
    private static bool CanWriteDirectory(string directory) { try { var p = Path.Combine(directory, $".plenaro-write-{Guid.NewGuid():N}"); using (File.Create(p)) { } File.Delete(p); return true; } catch { return false; } }
    private static string Read(JsonElement element, string name) => element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null ? value.GetString() ?? string.Empty : string.Empty;
    public void Dispose() { _shutdown.Cancel(); _shutdown.Dispose(); _client.Dispose(); _gate.Dispose(); }
}
