using Microsoft.OpenApi.Models;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "xls Converter API",
        Contact = new OpenApiContact
        {
            Email = "kirilg.ex@gmail.com"
        },
        License = new OpenApiLicense
        {
            Name = "MIT License"
        },
        Version = "v1"
    });

    options.IncludeXmlComments(
        Path.Combine(AppContext.BaseDirectory,
        $"{Assembly.GetExecutingAssembly().GetName().Name}.xml"));
});

var app = builder.Build();

EnsureTempFileDirectoryExists(app.Configuration);

app.UseSwagger();
app.UseSwaggerUI();

app.UseHttpsRedirection();

app.MapControllers();

app.Run();

void EnsureTempFileDirectoryExists(IConfiguration config)
{
    var path = Path.Combine(AppContext.BaseDirectory, config["TempDirectory"]);
    
    if (Directory.Exists(path) == false)
        Directory.CreateDirectory(path);
}
