using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using VB_build;

namespace VB_build.Api.Tests
{
    /// <summary>
    /// Boots the real API host but swaps in a known test API key and a <see cref="FakeBuildRunner"/>
    /// in place of the real VB6 build pipeline, so tests are fast/deterministic and don't require
    /// VB6 installed.
    /// </summary>
    public class ApiTestFactory : WebApplicationFactory<Program>
    {
        public const string TestApiKey = "test-api-key";

        public FakeBuildRunner BuildRunner { get; } = new FakeBuildRunner();

        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.ConfigureAppConfiguration((_, config) =>
            {
                config.AddInMemoryCollection(new Dictionary<string, string?>
                {
                    ["ApiKey"] = TestApiKey
                });
            });

            builder.ConfigureServices(services =>
            {
                services.RemoveAll<IBuildRunner>();
                services.AddSingleton<IBuildRunner>(BuildRunner);
            });
        }
    }
}
