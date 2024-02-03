using ExcelDna.Registration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace YahooXL;

public sealed class Options
{
    public int RtdIntervalMinSeconds { get; set; } = 3;
    public int RtdIntervalMaxSeconds { get; set; } = 60;
    public Options()
    {
        new ConfigurationBuilder()
            .AddEnvironmentVariables("YahooXL__")
            .Build()
            .Bind(this, opt => opt.ErrorOnUnknownConfiguration = true);
    }
}

public class AddIn : IExcelAddIn
{
    public static ILoggerFactory LogFactory { get; }
    public static ILogger Logger { get; }
    public static Options Options { get; } = new();

    static AddIn()
    {
        LogFactory = LoggerFactory
            .Create(builder => builder
                .AddDebug()
                .AddEventSourceLogger()
                .AddEventLog()
                .AddFilter("YahooQuotes", LogLevel.Information)
                .SetMinimumLevel(LogLevel.Debug));

        Logger = LogFactory.CreateLogger("YahooQuotesAddin");

        Logger.LogInformation("ExcelAddin: {Info}", Assembly.GetExecutingAssembly().GetName().ToString());
        //dynamic app = ExcelDnaUtil.Application;
        //var interval = app.RTD.ThrottleInterval; // default is 2s
    }

    public void AutoOpen()
    {
        // Since we have specified ExplicitRegistration=true in the .dna file, we need to do all registration explicitly.
        // Here we only add the async processing, which applies to our IObservable function.
        ExcelRegistration.GetExcelFunctions()
                         .ProcessAsyncRegistrations()
                         .RegisterFunctions();

        ExcelRegistration.GetExcelCommands().RegisterCommands();

        //string xllPath = (string)XlCall.Excel(XlCall.xlGetName);
        //var xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        //xlApp.AddIns.Add(xllPath, false /**don't copy file**/).Installed = true;
        Logger.LogTrace("Autoopen.");
    }

    public void AutoClose()
    {
        Logger.LogTrace("Autoclose.");
        YahooQuotesAddin.RefreshDisposable.Dispose();
    }
}
