using Microsoft.Extensions.Options;
using VB_build;
using VB_build.Api.Models;
using VB_build.Api.Security;
using VB_build.Api.Services;

var builder = WebApplication.CreateBuilder(args);

// Install JobConsoleWriter once at startup: it transparently passes all console output through
// to the original stdout, while also capturing per-job progress lines (see BuildQueueService)
// for the status-polling endpoint below. No changes needed in VB_build.Core for this to work.
Console.SetOut(new JobConsoleWriter(Console.Out));

// BuildOptions (VB6.EXE / verpatch / temp folder paths) are bound from configuration
// (appsettings.json "BuildOptions" section / environment variables), falling back to the
// same defaults the original hardcoded console app used.
builder.Services.Configure<BuildOptions>(builder.Configuration.GetSection("BuildOptions"));
builder.Services.Configure<ApiKeyOptions>(builder.Configuration);

builder.Services.AddSingleton<IBuildRunner>(sp =>
    new VbBuildRunner(sp.GetRequiredService<IOptions<BuildOptions>>().Value));

builder.Services.AddSingleton<BuildJobStore>();
builder.Services.AddSingleton<BuildJobQueue>();
builder.Services.AddHostedService<BuildQueueService>();

var app = builder.Build();

// POST /api/build/trigger - enqueues a build job and returns immediately with a job id.
// Jenkins should call this, then poll GET /api/build/status/{jobId} until Status is
// "Completed" or "Failed".
app.MapPost("/api/build/trigger", (BuildTriggerRequest request, BuildJobStore store, BuildJobQueue queue) =>
{
    var job = new BuildJob
    {
        Id = Guid.NewGuid().ToString("N"),
        Request = new BuildRequest
        {
            IsDll = request.IsDll,
            BuildPath = request.BuildPath,
            RefPath = request.RefPath,
            Version = request.Version,
            Compat = request.Compat,
            Time = request.Time,
            Command = request.Command
        }
    };

    store.Add(job);
    queue.Enqueue(job.Id);

    return Results.Accepted($"/api/build/status/{job.Id}", ToResponse(job));
})
.AddEndpointFilter<ApiKeyEndpointFilter>();

// GET /api/build/status/{jobId}?since=N - current status/result of a previously triggered
// build job. "since" (default 0) lets Jenkins do incremental polling: pass back the previous
// response's "nextSince" to receive only new progress lines since the last poll.
app.MapGet("/api/build/status/{jobId}", (string jobId, BuildJobStore store, int since = 0) =>
{
    if (!store.TryGet(jobId, out BuildJob job))
    {
        return Results.NotFound();
    }

    return Results.Ok(ToResponse(job, since));
})
.AddEndpointFilter<ApiKeyEndpointFilter>();

app.Run();

static BuildJobStatusResponse ToResponse(BuildJob job, int since = 0) => new()
{
    JobId = job.Id,
    Status = job.Status.ToString(),
    Success = job.Result?.Success,
    ErrorCount = job.Result?.ErrorCount,
    ErrorMessage = job.Result?.ErrorMessage,
    CreatedAt = job.CreatedAt,
    CompletedAt = job.CompletedAt,
    CurrentStep = job.CurrentStep,
    LogLines = job.GetLogLinesSince(since),
    NextSince = job.LogLineCount
};

// Exposes the generated Program class publicly so WebApplicationFactory<Program> can be used
// from the VB_build.Api.Tests integration test project.
public partial class Program { }
