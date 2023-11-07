using Microsoft.Extensions.Logging;

namespace YahooXL;

public static class Yahoo
{
    private static readonly ILogger Logger = LoggerFactory
        .Create(builder => builder
            .AddDebug()
            .AddEventSourceLogger()
#if DEBUG
            .SetMinimumLevel(LogLevel.Information))
#else
            .SetMinimumLevel(LogLevel.Warning))
#endif
        .CreateLogger("YahooFinance");

    private static readonly YahooQuotes YahooQuotes = new YahooQuotesBuilder()
            .WithLogger(Logger)
            .Build();

    public static async Task<Dictionary<string, Security?>> GetSecuritiesAsync(IEnumerable<string> symbols, CancellationToken ct)
    {
        return await YahooQuotes.GetAsync(symbols, Histories.None, "", ct).ConfigureAwait(false);
    }
}
