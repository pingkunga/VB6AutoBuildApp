using System.Net;
using System.Net.Http.Json;
using VB_build;
using VB_build.Api.Models;
using VB_build.Api.Security;
using Xunit;

namespace VB_build.Api.Tests
{
    public class BuildTriggerEndpointTests : IClassFixture<ApiTestFactory>
    {
        private readonly ApiTestFactory factory;

        public BuildTriggerEndpointTests(ApiTestFactory factory)
        {
            this.factory = factory;
        }

        private static BuildTriggerRequest SampleRequest() => new()
        {
            IsDll = true,
            BuildPath = @"C:\fake\build",
            RefPath = @"C:\fake\ref",
            Version = "1.0.0",
            Compat = "-n",
            Time = "1",
            Command = Array.Empty<string>()
        };

        [Fact]
        public async Task Trigger_WithoutApiKey_ReturnsUnauthorized()
        {
            HttpClient client = factory.CreateClient();

            HttpResponseMessage response = await client.PostAsJsonAsync("/api/build/trigger", SampleRequest());

            Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        }

        [Fact]
        public async Task Trigger_WithWrongApiKey_ReturnsUnauthorized()
        {
            HttpClient client = factory.CreateClient();
            client.DefaultRequestHeaders.Add(ApiKeyEndpointFilter.HeaderName, "wrong-key");

            HttpResponseMessage response = await client.PostAsJsonAsync("/api/build/trigger", SampleRequest());

            Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        }

        [Fact]
        public async Task Status_UnknownJobId_ReturnsNotFound()
        {
            HttpClient client = factory.CreateClient();
            client.DefaultRequestHeaders.Add(ApiKeyEndpointFilter.HeaderName, ApiTestFactory.TestApiKey);

            HttpResponseMessage response = await client.GetAsync("/api/build/status/does-not-exist");

            Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        }

        [Fact]
        public async Task Trigger_ThenStatus_CompletesUsingFakeRunner()
        {
            factory.BuildRunner.RunHandler = _ => new BuildResult { ErrorCount = 0, StepResults = new bool[10] };

            HttpClient client = factory.CreateClient();
            client.DefaultRequestHeaders.Add(ApiKeyEndpointFilter.HeaderName, ApiTestFactory.TestApiKey);

            HttpResponseMessage triggerResponse = await client.PostAsJsonAsync("/api/build/trigger", SampleRequest());
            Assert.Equal(HttpStatusCode.Accepted, triggerResponse.StatusCode);

            BuildJobStatusResponse? triggered = await triggerResponse.Content.ReadFromJsonAsync<BuildJobStatusResponse>();
            Assert.NotNull(triggered);
            Assert.False(string.IsNullOrEmpty(triggered!.JobId));

            BuildJobStatusResponse? final = null;
            for (int i = 0; i < 20; i++)
            {
                HttpResponseMessage statusResponse = await client.GetAsync($"/api/build/status/{triggered.JobId}");
                Assert.Equal(HttpStatusCode.OK, statusResponse.StatusCode);
                final = await statusResponse.Content.ReadFromJsonAsync<BuildJobStatusResponse>();

                if (final!.Status is "Completed" or "Failed")
                {
                    break;
                }

                await Task.Delay(50);
            }

            Assert.NotNull(final);
            Assert.Equal("Completed", final!.Status);
            Assert.True(final.Success);
            Assert.Single(factory.BuildRunner.Received);
            Assert.Equal(@"C:\fake\build", factory.BuildRunner.Received[0].BuildPath);
        }

        [Fact]
        public async Task Trigger_ThenStatus_ReportsFailureFromFakeRunner()
        {
            factory.BuildRunner.RunHandler = _ => new BuildResult
            {
                ErrorCount = 2,
                ErrorMessage = "simulated failure",
                StepResults = new bool[10]
            };

            HttpClient client = factory.CreateClient();
            client.DefaultRequestHeaders.Add(ApiKeyEndpointFilter.HeaderName, ApiTestFactory.TestApiKey);

            HttpResponseMessage triggerResponse = await client.PostAsJsonAsync("/api/build/trigger", SampleRequest());
            BuildJobStatusResponse? triggered = await triggerResponse.Content.ReadFromJsonAsync<BuildJobStatusResponse>();

            BuildJobStatusResponse? final = null;
            for (int i = 0; i < 20; i++)
            {
                HttpResponseMessage statusResponse = await client.GetAsync($"/api/build/status/{triggered!.JobId}");
                final = await statusResponse.Content.ReadFromJsonAsync<BuildJobStatusResponse>();

                if (final!.Status is "Completed" or "Failed")
                {
                    break;
                }

                await Task.Delay(50);
            }

            Assert.NotNull(final);
            Assert.Equal("Failed", final!.Status);
            Assert.False(final.Success);
            Assert.Equal("simulated failure", final.ErrorMessage);
        }
    }
}
