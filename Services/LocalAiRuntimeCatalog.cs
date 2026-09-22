using TaskTool.Models;

namespace TaskTool.Services;

public sealed record LocalAiRuntimeDefinition(
    string Version,
    string Backend,
    string FileName,
    string DownloadUrl,
    string Sha256,
    string License,
    string Platform);

public static class LocalAiRuntimeCatalog
{
    public static LocalAiRuntimeDefinition Cpu { get; } = new(
        Version: "b11081",
        Backend: "cpu",
        FileName: "llama-b11081-bin-win-cpu-x64.zip",
        DownloadUrl: "https://github.com/ggml-org/llama.cpp/releases/download/b11081/llama-b11081-bin-win-cpu-x64.zip",
        Sha256: "48f13c153946cca8543fd3ab915709ec5f340bfe1f38c687121b3d58f848b7b2",
        License: "MIT",
        Platform: "win-cpu-x64");

    public static LocalAiRuntimeDefinition Vulkan { get; } = new(
        Version: "b11081",
        Backend: "vulkan",
        FileName: "llama-b11081-bin-win-vulkan-x64.zip",
        DownloadUrl: "https://github.com/ggml-org/llama.cpp/releases/download/b11081/llama-b11081-bin-win-vulkan-x64.zip",
        Sha256: "4259a1dda3ef3fcfd8b007a16329d5bdcef07da8f5f95fddd85ff2954263f01a",
        License: "MIT",
        Platform: "win-vulkan-x64");

    public static LocalAiRuntimeDefinition Current => Cpu;
    public static LocalAiRuntimeDefinition Get(LocalAiComputeMode mode) => mode == LocalAiComputeMode.Gpu ? Vulkan : Cpu;
    public static string DirectoryName(LocalAiComputeMode mode) => mode == LocalAiComputeMode.Gpu ? "vulkan" : "cpu";
}

internal sealed record LocalAiRuntimeMetadata(
    string Provider,
    string Version,
    string Backend,
    string Platform,
    string ArchiveSha256);
