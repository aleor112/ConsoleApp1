using ConsoleApp1;
using log4net.Config;
using log4net;
using Microsoft.Extensions.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection.Emit;
using System.Threading.Tasks;
class Program
{

    static async Task Main()
    {
        // Load configuration from appsettings.json
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json")
            .Build();

        TestProject1.Const.urlPlaceholder = configuration["urlWithPlaceholder"];
        Console.WriteLine("My url is: " + TestProject1.Const.urlPlaceholder);
        // Configure log4net
        var log4netConfig = new FileInfo("log4net.config");
        XmlConfigurator.Configure(log4netConfig);

        // Your application code here

        // Example: Log an error
        ILog log = LogManager.GetLogger(typeof(Program));

        var generator = new ReportGenerator(log);
        generator.GenerateReport();


        await Console.Out.WriteLineAsync("Finished generating report");

    }

  
}