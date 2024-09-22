using PdfSharp.Fonts;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using Serilog;
using System;
using System.IO;
using ECMDocumentHelper.Helpers;

public class Program
{
    public static void Main(string[] args)
    {
        // Ensure Logs directory exists
        Directory.CreateDirectory("Logs");

        // Configure Serilog for file, console, and debug logging
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.Debug()
            .WriteTo.File("Logs/ECMDocumentHelper-.log", rollingInterval: RollingInterval.Day) // File logging
            .CreateLogger();

        // Set the font resolver globally
       GlobalFontSettings.FontResolver = new CustomFontResolver("Fonts"); // Replace "Fonts" with the actual directory where you stored the font file

        try
        {
            Log.Information("Starting ECMDocumentHelper application");
            CreateHostBuilder(args).Build().Run();
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Application startup failed");
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    public static IHostBuilder CreateHostBuilder(string[] args) =>
        Host.CreateDefaultBuilder(args)
            .UseSerilog() // Enable Serilog
            .ConfigureWebHostDefaults(webBuilder =>
            {
                webBuilder.UseStartup<Startup>();
            });
}
