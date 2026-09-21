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
    public void Settings_HaveSafeAiDefaults()
    {
        var settings = new AppSettings();
        Assert.False(settings.AiEnabled);
        Assert.Equal(AiProviderType.OpenAiCompatible, settings.AiProvider);
        Assert.Equal(LocalAiPreset.Light, settings.AiLocalPreset);
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
    public async Task HashValidation_RejectsMismatchAndPartIsNotAValidModel()
    {
        var file = Path.GetTempFileName();
        try
        {
            await File.WriteAllTextAsync(file, "not a model");
            Assert.False(await LocalLlamaServerManager.HasSha256Async(file, new string('0', 64)));
            Assert.EndsWith(".part", file + ".part", StringComparison.Ordinal);
        }
        finally { File.Delete(file); }
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

    private static HttpResponseMessage Json(HttpStatusCode status, string body)
        => new(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };

    private sealed class StubHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> response) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => response(request);
    }
}
