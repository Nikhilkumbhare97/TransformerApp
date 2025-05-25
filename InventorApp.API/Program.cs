using InventorApp.API.Services;
using InventorApp.API.Repositories;
using InventorApp.API.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Localization;
using System.Globalization;
using Pomelo.EntityFrameworkCore.MySql.Infrastructure;
using Npgsql.EntityFrameworkCore.PostgreSQL.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Add globalization configuration
builder.Services.Configure<RequestLocalizationOptions>(options =>
{
    options.DefaultRequestCulture = new Microsoft.AspNetCore.Localization.RequestCulture("en-US");
    options.SupportedCultures = new List<CultureInfo> { new CultureInfo("en-US") };
    options.SupportedUICultures = new List<CultureInfo> { new CultureInfo("en-US") };
});

// Add CORS configuration
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigin",
        builder =>
        {
            builder.AllowAnyOrigin()
                   .AllowAnyHeader()
                   .AllowAnyMethod();
        });
});

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register ProjectService and IProjectRepository
builder.Services.AddScoped<ProjectService>();
builder.Services.AddScoped<IProjectRepository, ProjectRepository>();

// Register TransformerService and ITransformerRepository
builder.Services.AddScoped<TransformerService>();
builder.Services.AddScoped<ITransformerRepository, TransformerRepository>();

// Register TransformerConfigurationService and Repository
builder.Services.AddScoped<TransformerConfigurationService>();
builder.Services.AddScoped<ITransformerConfigurationRepository, TransformerConfigurationRepository>();

// Register ImageConfig services
builder.Services.AddScoped<ImageConfigService>();
builder.Services.AddScoped<IImageConfigRepository, ImageConfigRepository>();

// Register Assembly services
builder.Services.AddSingleton<AssemblyService>();

// Add DbContext configuration
builder.Services.AddDbContext<ApplicationDbContext>((serviceProvider, options) =>
{
    var configuration = serviceProvider.GetRequiredService<IConfiguration>();
    var dbType = configuration.GetValue<string>("Database:Type", "PostgreSQL"); // Default to PostgreSQL

    if (dbType.Equals("PostgreSQL", StringComparison.OrdinalIgnoreCase))
    {
        options.UseNpgsql(
            configuration.GetConnectionString("PostgresConnection"),
            npgsqlOptions => npgsqlOptions.EnableRetryOnFailure()
        );
    }
    else if (dbType.Equals("MySQL", StringComparison.OrdinalIgnoreCase))
    {
        options.UseMySql(
            configuration.GetConnectionString("MySqlConnection"),
            ServerVersion.AutoDetect(configuration.GetConnectionString("MySqlConnection")),
            mySqlOptions => mySqlOptions.EnableRetryOnFailure()
        );
    }
    else
    {
        throw new InvalidOperationException($"Unsupported database type: {dbType}");
    }
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseRequestLocalization();
app.UseHttpsRedirection();
app.UseRouting();
app.UseCors("AllowSpecificOrigin");
app.UseAuthorization();
app.MapControllers();

// Add this section before app.Run()
using (var scope = app.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    try
    {
        var context = services.GetRequiredService<ApplicationDbContext>();
        context.Database.EnsureCreated();
    }
    catch (Exception ex)
    {
        var logger = services.GetRequiredService<ILogger<Program>>();
        logger.LogError(ex, "An error occurred while creating the database.");
    }
}

app.Run();