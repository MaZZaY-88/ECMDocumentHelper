using ECMDocumentHelper.Helpers;
using ECMDocumentHelper.Services;
using Microsoft.Extensions.DependencyInjection;
using PdfSharp.Fonts;

var builder = WebApplication.CreateBuilder(args);


// Add controllers
// Add services to the container.
builder.Services.AddControllers();

// Register OfficeInteropHelper
builder.Services.AddScoped<OfficeInteropHelper>();

// Register PdfProcessingService
builder.Services.AddScoped<PdfProcessingService>();

// Set the font resolver globally
GlobalFontSettings.FontResolver = new CustomFontResolver("Fonts");

var app = builder.Build();

// Configure HTTP request pipeline
app.UseRouting();
app.MapControllers();

app.Run();
