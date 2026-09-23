using System.IO;
using System.Threading;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class AiKnowledgeService : IDisposable
{
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private FileSystemWatcher? _watcher;
    private readonly Timer _debounce;
    public AiKnowledgeIndexService Index { get; }
    public AiKnowledgeSearchService Search { get; }
    public event EventHandler? StatusChanged;
    public AiKnowledgeIndexStatus Status { get; private set; } = new(0, 0, null);

    public AiKnowledgeService(SettingsService settings, LoggerService logger, AiKnowledgeIndexService? index = null)
    {
        _settings = settings; _logger = logger; Index = index ?? new AiKnowledgeIndexService(logger); Search = new(Index.IndexPath, logger);
        _debounce = new Timer(async _ => await IndexSafelyAsync(false), null, Timeout.Infinite, Timeout.Infinite);
        if (settings.Current.AiKnowledgeEnabled) SetEnabled(true);
    }
    public void SetEnabled(bool enabled)
    {
        if (!enabled) { if (_watcher != null) _watcher.EnableRaisingEvents = false; return; }
        Index.EnsureKnowledgeDirectory();
        if (_watcher == null)
        {
            _watcher = new FileSystemWatcher(Index.KnowledgePath) { IncludeSubdirectories = true };
            _watcher.Created += OnChanged; _watcher.Changed += OnChanged; _watcher.Deleted += OnChanged; _watcher.Renamed += OnChanged;
            _watcher.Error += (_, e) => _logger.Warning($"[AI Knowledge] Watcher error='{e.GetException().Message}'");
        }
        _watcher.EnableRaisingEvents = true; _ = IndexSafelyAsync(false);
    }
    public Task RebuildAsync() => IndexSafelyAsync(true);
    public async Task<IReadOnlyList<AiKnowledgeMatch>> SearchAsync(string question, CancellationToken ct = default)
        => !_settings.Current.AiKnowledgeEnabled ? Array.Empty<AiKnowledgeMatch>() : await Search.SearchAsync(question, ct: ct);
    private void OnChanged(object sender, FileSystemEventArgs e) { if (_settings.Current.AiKnowledgeEnabled) _debounce.Change(TimeSpan.FromSeconds(3), Timeout.InfiniteTimeSpan); }
    private async Task IndexSafelyAsync(bool rebuild)
    {
        if (!_settings.Current.AiKnowledgeEnabled && !rebuild) return;
        try { Status = await Index.IndexAsync(rebuild); }
        catch (Exception ex) { Status = Status with { Error = ex.Message }; _logger.Warning($"[AI Knowledge] Index failed error='{ex.Message}'"); }
        StatusChanged?.Invoke(this, EventArgs.Empty);
    }
    public void Dispose() { _watcher?.Dispose(); _debounce.Dispose(); }
}
