namespace YahooXL;

internal static class YahooFinance
{
    private static readonly YahooQuotes YahooQuotes = new YahooQuotesBuilder()
        .WithLogger(AddIn.LogFactory.CreateLogger("YahooQuotes"))
        .Build();

    internal static async Task<Dictionary<string, Security?>> GetSecuritiesAsync(IEnumerable<string> symbols, CancellationToken ct) =>
        await YahooQuotes.GetAsync(symbols, Histories.None, "", ct).ConfigureAwait(false);
}
