namespace TaskTool.Models;

public enum AiChatRole { System, User, Assistant }

public sealed record AiChatRequestMessage(AiChatRole Role, string Content);

public sealed record AiRequestOptions(double Temperature = 0.3, int MaxTokens = 1024);

public sealed record AiKnowledgeSource(string RelativePath, int? PageNumber = null);

public sealed record AiChatMessage(AiChatRole Role, string Content, DateTime CreatedAt, IReadOnlyList<AiKnowledgeSource>? KnowledgeSources = null)
{
    public string Author => Role == AiChatRole.User ? "Du" : "Plenaro KI";
    public bool IsUser => Role == AiChatRole.User;
    public bool IsAssistant => Role == AiChatRole.Assistant;
    public bool HasKnowledgeSources => KnowledgeSources?.Count > 0;
    public string KnowledgeSummary => HasKnowledgeSources
        ? $"Lokales Wissen: {KnowledgeSources!.Count} Quelle{(KnowledgeSources.Count == 1 ? string.Empty : "n")} · {string.Join(", ", KnowledgeSources.Select(x => x.RelativePath))}"
        : string.Empty;
}
