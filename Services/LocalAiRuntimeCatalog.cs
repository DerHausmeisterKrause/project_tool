namespace TaskTool.Services;

public sealed record LocalAiRuntimeDefinition(
    string Version,
    string FileName,
    string DownloadUrl,
    string Sha256,
    string License,
    string Platform);

public static class LocalAiRuntimeCatalog
{
    public static LocalAiRuntimeDefinition Current { get; } = new(
        Version: "b11081",
        FileName: "llama-b11081-bin-win-cpu-x64.zip",
        DownloadUrl: "https://github.com/ggml-org/llama.cpp/releases/download/b11081/llama-b11081-bin-win-cpu-x64.zip",
        Sha256: "48f13c153946cca8543fd3ab915709ec5f340bfe1f38c687121b3d58f848b7b2",
        License: "MIT",
        Platform: "win-cpu-x64");
}

internal sealed record LocalAiRuntimeMetadata(
    string Provider,
    string Version,
    string Platform,
    string ArchiveSha256);
