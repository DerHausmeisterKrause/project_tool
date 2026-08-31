using System.Net;

namespace TaskTool.Models;

public sealed class ZnunyApiException : Exception
{
    public string Stage { get; }
    public HttpStatusCode StatusCode { get; }
    public string ErrorCode { get; }
    public string ErrorMessage { get; }
    public string ResponseBody { get; }

    public ZnunyApiException(string stage, HttpStatusCode statusCode, string errorCode, string errorMessage, string responseBody)
        : base(string.IsNullOrWhiteSpace(errorCode) ? errorMessage : $"{errorCode}: {errorMessage}")
    {
        Stage = stage;
        StatusCode = statusCode;
        ErrorCode = errorCode;
        ErrorMessage = errorMessage;
        ResponseBody = responseBody;
    }
}
