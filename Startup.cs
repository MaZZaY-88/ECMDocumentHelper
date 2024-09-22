using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ECMDocumentHelper.Helpers;
using ECMDocumentHelper.Services;

public class Startup
{
    public Startup(IConfiguration configuration)
    {
        Configuration = configuration;
    }

    public IConfiguration Configuration { get; }

    public void ConfigureServices(IServiceCollection services)
    {
        services.AddControllers();

        // Inject configuration values for output directories from appsettings.json
        var outputDirectory = Configuration["PdfSettings:outputDirectory"];
        var pdfSaveDirectory = Configuration["PdfSettings:pdfSaveDirectory"];
        var barcodeProfile = Configuration.GetSection("BarcodeProfile");

        services.AddScoped<OfficeInteropHelper>();

        // Register ImageHelper and BarcodeHelper with required configuration values
        services.AddScoped<ImageHelper>(provider =>
            new ImageHelper(outputDirectory, provider.GetRequiredService<ILogger<ImageHelper>>())
        );

        services.AddSingleton<BarcodeHelper>(provider =>
            new BarcodeHelper(
            provider.GetRequiredService<IConfiguration>(),
            provider.GetRequiredService<ILogger<BarcodeHelper>>()
        ));


        services.AddScoped<PdfProcessingService>();
    }

    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        app.UseRouting();

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllers();
        });
    }
}
