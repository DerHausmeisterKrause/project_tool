namespace TaskTool.Models;

public enum UpdateState { Idle, Checking, UpToDate, UpdateAvailable, Downloading, ReadyToInstall, Installing, Failed }

public sealed record GitHubReleaseAsset(string Name, string DownloadUrl, long Size, string Digest);

public sealed record UpdateInfo(
    Version Version,
    string TagName,
    string Name,
    string ReleaseNotes,
    string HtmlUrl,
    DateTimeOffset? PublishedAt,
    GitHubReleaseAsset Asset);

public sealed record UpdateCheckResult(bool UpdateAvailable, UpdateInfo? Update, string Message);
