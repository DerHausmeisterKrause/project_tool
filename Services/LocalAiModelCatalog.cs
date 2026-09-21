using TaskTool.Models;

namespace TaskTool.Services;

public sealed record LocalAiModelDefinition(
    LocalAiPreset Id, string DisplayName, string Repository, string FileName,
    string DownloadUrl, string Sha256, string ApproximateSize, string License, string LlamaAlias);

public static class LocalAiModelCatalog
{
    public static IReadOnlyList<LocalAiModelDefinition> All { get; } =
    [
        new(LocalAiPreset.Light, "Light – Qwen3 1.7B – ca. 1,3 GB", "ggml-org/Qwen3-1.7B-GGUF", "Qwen3-1.7B-Q4_K_M.gguf", "https://huggingface.co/ggml-org/Qwen3-1.7B-GGUF/resolve/main/Qwen3-1.7B-Q4_K_M.gguf", "d2387ca2dbfee2ffabce7120d3770dadca0b293052bc2f0e138fdc940d9bc7b5", "ca. 1,3 GB", "Apache-2.0", "plenaro-local"),
        new(LocalAiPreset.Middle, "Middle – SmolLM3 3B – ca. 1,9 GB", "ggml-org/SmolLM3-3B-GGUF", "SmolLM3-Q4_K_M.gguf", "https://huggingface.co/ggml-org/SmolLM3-3B-GGUF/resolve/main/SmolLM3-Q4_K_M.gguf", "8334b850b7bd46238c16b0c550df2138f0889bf433809008cc17a8b05761863e", "ca. 1,9 GB", "Apache-2.0", "plenaro-local"),
        new(LocalAiPreset.High, "High – Qwen3 8B – ca. 5,0 GB", "Qwen/Qwen3-8B-GGUF", "Qwen3-8B-Q4_K_M.gguf", "https://huggingface.co/Qwen/Qwen3-8B-GGUF/resolve/main/Qwen3-8B-Q4_K_M.gguf", "d98cdcbd03e17ce47681435b5150e34c1417f50b5c0019dd560e4882c5745785", "ca. 5,0 GB", "Apache-2.0", "plenaro-local")
    ];

    public static LocalAiModelDefinition Get(LocalAiPreset preset) =>
        All.Single(model => model.Id == preset);
}
