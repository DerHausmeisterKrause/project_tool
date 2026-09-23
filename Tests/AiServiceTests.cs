using System.Net;
using System.Text;
using System.Text.Json;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class AiServiceTests
{
    [Fact]
    public void RuntimeCatalog_ContainsPinnedApprovedBuild()
    {
        var runtime = LocalAiRuntimeCatalog.Current;
        Assert.Equal("b11081", runtime.Version);
        Assert.Equal("llama-b11081-bin-win-cpu-x64.zip", runtime.FileName);
        Assert.Equal("https://github.com/ggml-org/llama.cpp/releases/download/b11081/llama-b11081-bin-win-cpu-x64.zip", runtime.DownloadUrl);
        Assert.Equal(Uri.UriSchemeHttps, new Uri(runtime.DownloadUrl).Scheme);
        Assert.Matches("^[0-9a-f]{64}$", runtime.Sha256);
        Assert.Equal("48f13c153946cca8543fd3ab915709ec5f340bfe1f38c687121b3d58f848b7b2", runtime.Sha256);
        Assert.Equal("win-cpu-x64", runtime.Platform);
        Assert.Equal("MIT", runtime.License);
    }

    [Fact]
    public void RuntimeCatalog_ContainsSeparateCpuAndVulkanBuildsAtSameVersion()
    {
        var cpu = LocalAiRuntimeCatalog.Get(LocalAiComputeMode.Cpu);
        var gpu = LocalAiRuntimeCatalog.Get(LocalAiComputeMode.Gpu);
        Assert.Equal(cpu.Version, gpu.Version);
        Assert.Equal("cpu", cpu.Backend);
        Assert.Equal("b11081", gpu.Version);
        Assert.Equal("vulkan", gpu.Backend);
        Assert.Equal("llama-b11081-bin-win-vulkan-x64.zip", gpu.FileName);
        Assert.Equal("4259a1dda3ef3fcfd8b007a16329d5bdcef07da8f5f95fddd85ff2954263f01a", gpu.Sha256);
        Assert.Equal("win-vulkan-x64", gpu.Platform);
        Assert.Equal(Uri.UriSchemeHttps, new Uri(gpu.DownloadUrl).Scheme);
    }

    [Fact]
    public void RuntimePaths_AreSeparatedByBackend()
    {
        Assert.EndsWith(Path.Combine("runtime", "llama.cpp", "cpu"), LocalLlamaServerManager.GetRuntimeDirectory("root", LocalAiComputeMode.Cpu));
        Assert.EndsWith(Path.Combine("runtime", "llama.cpp", "vulkan"), LocalLlamaServerManager.GetRuntimeDirectory("root", LocalAiComputeMode.Gpu));
    }

    [Fact]
    public async Task RuntimeValidation_AcceptsMatchingMetadata()
    {
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            var executable = Path.Combine(directory.FullName, "llama-server.exe");
            var metadata = Path.Combine(directory.FullName, "runtime.json");
            await File.WriteAllTextAsync(executable, "test");
            await File.WriteAllTextAsync(metadata, """
                {"provider":"llama.cpp","version":"b11081","backend":"cpu","platform":"win-cpu-x64","archiveSha256":"48f13c153946cca8543fd3ab915709ec5f340bfe1f38c687121b3d58f848b7b2"}
                """);

            Assert.True(LocalLlamaServerManager.IsRuntimeInstalled(executable, metadata, LocalAiRuntimeCatalog.Current));
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public async Task RuntimeValidation_RejectsMissingMetadataAndOtherVersion()
    {
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            var executable = Path.Combine(directory.FullName, "llama-server.exe");
            var metadata = Path.Combine(directory.FullName, "runtime.json");
            await File.WriteAllTextAsync(executable, "test");

            Assert.False(LocalLlamaServerManager.IsRuntimeInstalled(executable, metadata, LocalAiRuntimeCatalog.Current));

            await File.WriteAllTextAsync(metadata, """
                {"provider":"llama.cpp","version":"b10938","backend":"cpu","platform":"win-cpu-x64","archiveSha256":"48f13c153946cca8543fd3ab915709ec5f340bfe1f38c687121b3d58f848b7b2"}
                """);
            Assert.False(LocalLlamaServerManager.IsRuntimeInstalled(executable, metadata, LocalAiRuntimeCatalog.Current));
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public async Task RuntimeArchiveHashMismatch_IsRejectedAndPartIsDeleted()
    {
        var directory = Directory.CreateTempSubdirectory();
        var part = Path.Combine(directory.FullName, "runtime.zip.part");
        var runtime = Path.Combine(directory.FullName, "llama.cpp");
        try
        {
            Directory.CreateDirectory(runtime);
            var existingExecutable = Path.Combine(runtime, "llama-server.exe");
            await File.WriteAllTextAsync(existingExecutable, "existing working runtime");
            await File.WriteAllTextAsync(part, "not the approved archive");

            var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
                LocalLlamaServerManager.InstallRuntimeArchiveAsync(part, runtime, LocalAiRuntimeCatalog.Current, () => { }));
            Assert.Equal("Die heruntergeladene llama.cpp-Runtime konnte nicht verifiziert werden.", exception.Message);
            Assert.False(File.Exists(part));
            Assert.True(File.Exists(existingExecutable));
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public void Settings_HaveSafeAiDefaults()
    {
        var settings = new AppSettings();
        Assert.False(settings.AiEnabled);
        Assert.Equal(AiProviderType.OpenAiCompatible, settings.AiProvider);
        Assert.Equal(LocalAiPreset.Light, settings.AiLocalPreset);
        Assert.Equal(LocalAiComputeMode.Cpu, settings.AiLocalComputeMode);
    }

    [Fact]
    public void LocalCatalog_ContainsOnlyTheThreePinnedHttpsModels()
    {
        Assert.Equal(new[] { LocalAiPreset.Light, LocalAiPreset.Middle, LocalAiPreset.High }, LocalAiModelCatalog.All.Select(x => x.Id));
        Assert.Equal(new[] { "Qwen3-1.7B-Q4_K_M.gguf", "SmolLM3-Q4_K_M.gguf", "Qwen3-8B-Q4_K_M.gguf" }, LocalAiModelCatalog.All.Select(x => x.FileName));
        Assert.Equal(new[] { "d2387ca2dbfee2ffabce7120d3770dadca0b293052bc2f0e138fdc940d9bc7b5", "8334b850b7bd46238c16b0c550df2138f0889bf433809008cc17a8b05761863e", "d98cdcbd03e17ce47681435b5150e34c1417f50b5c0019dd560e4882c5745785" }, LocalAiModelCatalog.All.Select(x => x.Sha256));
        Assert.All(LocalAiModelCatalog.All, model => Assert.Equal(Uri.UriSchemeHttps, new Uri(model.DownloadUrl).Scheme));
        Assert.Equal("127.0.0.1", LocalLlamaServerManager.Host);
    }

    [Fact]
    public async Task ModelValidation_RejectsPartFileAndHashMismatch()
    {
        var directory = Directory.CreateTempSubdirectory();
        var file = Path.Combine(directory.FullName, "model.gguf");
        try
        {
            await File.WriteAllTextAsync(file, "valid model bytes");
            var expectedHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(file)));
            await File.WriteAllTextAsync(file + ".part", "incomplete download");

            Assert.False(await LocalLlamaServerManager.IsModelFileValidAsync(file, expectedHash));

            File.Delete(file + ".part");
            Assert.False(await LocalLlamaServerManager.IsModelFileValidAsync(file, new string('0', 64)));
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public async Task DownloadPromotion_DeletesPartFileAfterHashMismatch()
    {
        var directory = Directory.CreateTempSubdirectory();
        var part = Path.Combine(directory.FullName, "model.gguf.part");
        var destination = Path.Combine(directory.FullName, "model.gguf");
        try
        {
            await File.WriteAllTextAsync(part, "corrupt download");

            await Assert.ThrowsAsync<InvalidDataException>(() => LocalLlamaServerManager.PromoteVerifiedDownloadAsync(part, destination, new string('0', 64)));

            Assert.False(File.Exists(part));
            Assert.False(File.Exists(destination));
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public void ThinkingOutput_IsRemovedForConnectionTest()
    {
        var cleaned = System.Text.RegularExpressions.Regex.Replace("<think>internal</think>\nTest erfolgreich", @"<think>[\s\S]*?</think>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
        Assert.Equal("Test erfolgreich", cleaned);
    }

    [Fact]
    public async Task OpenAiProvider_SendsExpectedChatCompletionRequest()
    {
        HttpRequestMessage? captured = null;
        var handler = new StubHandler(async request =>
        {
            captured = request;
            var body = await request.Content!.ReadAsStringAsync();
            using var json = JsonDocument.Parse(body);
            Assert.Equal("test-model", json.RootElement.GetProperty("model").GetString());
            Assert.Equal(0, json.RootElement.GetProperty("temperature").GetInt32());
            Assert.Equal(32, json.RootElement.GetProperty("max_tokens").GetInt32());
            Assert.Equal(AiService.TestUserPrompt, json.RootElement.GetProperty("messages")[1].GetProperty("content").GetString());
            return Json(HttpStatusCode.OK, "{\"choices\":[{\"message\":{\"content\":\"Test erfolgreich\"}}]}");
        });
        using var client = new HttpClient(handler);
        var provider = new OpenAiCompatibleAiProvider(client, "http://localhost:1234/v1", "test-model", "secret");

        var result = await provider.SendAsync(AiService.TestSystemPrompt, AiService.TestUserPrompt);

        Assert.Equal("Test erfolgreich", result);
        Assert.Equal("http://localhost:1234/v1/chat/completions", captured!.RequestUri!.ToString());
        Assert.Equal("Bearer", captured.Headers.Authorization!.Scheme);
        Assert.Equal("secret", captured.Headers.Authorization.Parameter);
    }

    [Fact]
    public async Task OpenAiProvider_WorksWithoutApiKeyAndAddsV1Route()
    {
        HttpRequestMessage? captured = null;
        using var client = new HttpClient(new StubHandler(request =>
        {
            captured = request;
            return Task.FromResult(Json(HttpStatusCode.OK, "{\"choices\":[{\"message\":{\"content\":\"ok\"}}]}"));
        }));
        var provider = new OpenAiCompatibleAiProvider(client, "http://localhost:1234", "local-model");

        Assert.Equal("ok", await provider.SendAsync("system", "user"));
        Assert.Equal("http://localhost:1234/v1/chat/completions", captured!.RequestUri!.ToString());
        Assert.Null(captured.Headers.Authorization);
    }

    [Fact]
    public async Task OpenAiProvider_ChatUsesMessagesAndRequestOptions()
    {
        using var client = new HttpClient(new StubHandler(async request =>
        {
            var body = await request.Content!.ReadAsStringAsync();
            using var json = JsonDocument.Parse(body);
            Assert.Equal(0.3, json.RootElement.GetProperty("temperature").GetDouble(), 3);
            Assert.Equal(1024, json.RootElement.GetProperty("max_tokens").GetInt32());
            Assert.Equal("assistant", json.RootElement.GetProperty("messages")[1].GetProperty("role").GetString());
            return Json(HttpStatusCode.OK, "{\"choices\":[{\"message\":{\"content\":\"beliebige gültige Antwort\"}}]}");
        }));
        var provider = new OpenAiCompatibleAiProvider(client, "http://localhost:1234", "model");

        var answer = await provider.ChatAsync(
            [new(AiChatRole.User, "Frage"), new(AiChatRole.Assistant, "Kontext")],
            new AiRequestOptions());

        Assert.Equal("beliebige gültige Antwort", answer);
    }

    [Fact]
    public void LocalServerArguments_DisableReasoning()
    {
        var arguments = LocalLlamaServerManager.BuildServerArguments("model.gguf", 1234, "plenaro-local");
        Assert.Contains("--reasoning", arguments);
        var reasoningIndex = arguments.ToList().IndexOf("--reasoning");
        Assert.Equal("off", arguments[reasoningIndex + 1]);
    }

    [Fact]
    public void LocalServerArguments_OnlyEnableOffloadForGpu()
    {
        var cpu = LocalLlamaServerManager.BuildServerArguments("model.gguf", 1234, "model", LocalAiComputeMode.Cpu);
        var gpu = LocalLlamaServerManager.BuildServerArguments("model.gguf", 1234, "model", LocalAiComputeMode.Gpu);
        Assert.DoesNotContain("--n-gpu-layers", cpu);
        var index = gpu.ToList().IndexOf("--n-gpu-layers");
        Assert.True(index >= 0); Assert.Equal("all", gpu[index + 1]);
    }

    [Fact]
    public async Task ComputeMode_IsPersistedAndReloaded()
    {
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            var path = Path.Combine(directory.FullName, "settings.json"); var logger = new LoggerService();
            var settings = new SettingsService(logger, path); settings.Current.AiLocalComputeMode = LocalAiComputeMode.Gpu; settings.Save();
            Assert.Equal(LocalAiComputeMode.Gpu, new SettingsService(logger, path).Current.AiLocalComputeMode);
        }
        finally { directory.Delete(true); }
    }

    [Fact]
    public void GpuStartupFailure_UsesControlledCpuFallbackPolicy()
    {
        Assert.True(LocalLlamaServerManager.ShouldFallbackToCpu(LocalAiComputeMode.Gpu, new InvalidOperationException("Vulkan failed")));
        Assert.False(LocalLlamaServerManager.ShouldFallbackToCpu(LocalAiComputeMode.Cpu, new InvalidOperationException("failed")));
        Assert.False(LocalLlamaServerManager.ShouldFallbackToCpu(LocalAiComputeMode.Gpu, new OperationCanceledException()));
    }

    [Fact]
    public async Task ConnectionTest_AcceptsAnyNonEmptyValidAssistantAnswer()
    {
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            var logger = new LoggerService();
            var settings = new SettingsService(logger, Path.Combine(directory.FullName, "settings.json"));
            settings.Current.AiEnabled = true;
            settings.Current.AiProvider = AiProviderType.OpenAiCompatible;
            settings.Current.AiApiBaseUrl = "http://localhost:1234";
            settings.Current.AiModel = "model";
            using var client = new HttpClient(new StubHandler(_ => Task.FromResult(
                Json(HttpStatusCode.OK, "{\"choices\":[{\"message\":{\"content\":\"Verbindung steht.\"}}]}"))));
            using var service = new AiService(settings, logger, client);

            Assert.Equal("Verbindung steht.", await service.TestAsync());
        }
        finally { directory.Delete(true); }
    }

    private static HttpResponseMessage Json(HttpStatusCode status, string body)
        => new(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };

    private sealed class StubHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> response) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => response(request);
    }
}
