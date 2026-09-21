namespace TaskTool.Models;

public enum AiChatRole { System, User, Assistant }

public sealed record AiChatRequestMessage(AiChatRole Role, string Content);

public sealed record AiRequestOptions(double Temperature = 0.3, int MaxTokens = 1024);

public sealed record AiChatMessage(AiChatRole Role, string Content, DateTime CreatedAt)
{
    public string Author => Role == AiChatRole.User ? "Du" : "Plenaro KI";
    public bool IsUser => Role == AiChatRole.User;
}
