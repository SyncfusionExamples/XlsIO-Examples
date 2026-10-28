#nullable enable

using Microsoft.AspNetCore.Http.Features;
using Npgsql;
using Syncfusion.Licensing;
using AgentRAGWithAzure.Endpoints;
using AgentRAGWithAzure.Services;

var builder = WebApplication.CreateBuilder(args);

// ------------------------------------------------------------
// Configure PostgreSQL and pgvector
// ------------------------------------------------------------

string postgresConnectionString =
    builder.Configuration.GetConnectionString("PostgreSql")
    ?? throw new InvalidOperationException(
        "ConnectionStrings:PostgreSql is missing.");

builder.Services.AddSingleton<NpgsqlDataSource>(_ =>
{
    var dataSourceBuilder =
        new NpgsqlDataSourceBuilder(postgresConnectionString);

    // Enables the Pgvector.Vector type in Npgsql.
    dataSourceBuilder.UseVector();

    return dataSourceBuilder.Build();
});

builder.Services.AddSingleton<RagService>();
builder.Services.AddSingleton<UploadQuotaService>();
builder.Services.AddMemoryCache();
builder.Services.AddControllers();

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit =
        200 * 1024 * 1024;
});

var app = builder.Build();

// Register Syncfusion license for server-side DocIO / WordEditor (SFDT).
// Prefer Syncfusion:LicenseKey from appsettings; keep a local fallback for demos.
string? syncfusionLicense =
    builder.Configuration["Syncfusion:LicenseKey"];

if (string.IsNullOrWhiteSpace(syncfusionLicense))
{
    syncfusionLicense =
        "Ngo9BigBOggjHTQxAR8/V1JAaF5cXGhAd1p/WGJbYl1xZFdFallYTnNWSV9jS3hTc0RmWHlaeXVQQWdYUE91Vw==";
}

SyncfusionLicenseProvider.RegisterLicense(syncfusionLicense);

// ------------------------------------------------------------
// Block direct HTTP access to local RAG uploads
// ------------------------------------------------------------

app.Use(async (context, next) =>
{
    if (context.Request.Path.StartsWithSegments(
            "/rag/uploads",
            StringComparison.OrdinalIgnoreCase) ||
        context.Request.Path.StartsWithSegments(
            "/rag/store",
            StringComparison.OrdinalIgnoreCase))
    {
        context.Response.StatusCode =
            StatusCodes.Status403Forbidden;

        await context.Response.WriteAsync("Forbidden");

        return;
    }

    await next();
});

app.UseDefaultFiles();
app.UseStaticFiles();
app.MapControllers();

app.MapRagEndpoints();

app.Run();
