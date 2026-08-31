namespace TaskTool.Models;

internal static class WikiSearchPolicy
{
    internal static bool CanSearch(TaskItem? task, bool force)
        => task?.IsZnunyTask == true && (force || task.IsZnunyAssigned);

    internal static string ResolveTitle(TaskItem task, string? cachedTitle)
        => !string.IsNullOrWhiteSpace(cachedTitle) ? cachedTitle : task.Title;

    internal static string ResolveMessage(TaskItem task, string? cachedMessage)
        => cachedMessage ?? task.Description ?? string.Empty;
}
