using System.Net;
using System.Text;
using System.Text.Json;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class AiServiceTests
{
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
