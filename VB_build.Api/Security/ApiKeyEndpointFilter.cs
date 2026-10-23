using Microsoft.Extensions.Options;

namespace VB_build.Api.Security
{
    /// <summary>Configuration binding for the API key required on <c>/api/build/*</c> endpoints.</summary>
    public class ApiKeyOptions
    {
        public string ApiKey { get; set; } = string.Empty;
    }

    /// <summary>
    /// Minimal-API endpoint filter that requires a matching <c>X-Api-Key</c> header. This is the
    /// only gate on who can trigger a build on the machine, so a real, non-empty key must be
    /// configured (see appsettings.json / environment variable "ApiKey").
    /// </summary>
    public class ApiKeyEndpointFilter : IEndpointFilter
    {
        public const string HeaderName = "X-Api-Key";

        private readonly IOptionsMonitor<ApiKeyOptions> options;

        public ApiKeyEndpointFilter(IOptionsMonitor<ApiKeyOptions> options)
        {
            this.options = options;
        }

        public async ValueTask<object?> InvokeAsync(EndpointFilterInvocationContext context, EndpointFilterDelegate next)
        {
            string configuredKey = options.CurrentValue.ApiKey;

            if (string.IsNullOrEmpty(configuredKey))
            {
                return Results.Problem("Server API key is not configured.", statusCode: StatusCodes.Status500InternalServerError);
            }

            if (!context.HttpContext.Request.Headers.TryGetValue(HeaderName, out Microsoft.Extensions.Primitives.StringValues providedKey)
                || providedKey.Count != 1
                || providedKey[0] != configuredKey)
            {
                return Results.Unauthorized();
            }

            return await next(context);
        }
    }
}
