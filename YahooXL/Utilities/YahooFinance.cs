namespace YahooXL;

public class YahooQuotesData
{
    private readonly YahooQuotes YahooQuotes;

    public YahooQuotesData()
    {
        YahooQuotes = new YahooQuotesBuilder()
            //.WithLogger(logger)
            .WithCacheDuration(Duration.FromMinutes(30), Duration.FromHours(6))
            .WithHistoryStartDate(Instant.FromUtc(1990, 1, 1, 0, 0))
            .Build();
    }

    public async Task<Dictionary<string, Security?>> GetSecuritiesAsync(IEnumerable<string> symbols, CancellationToken ct) =>
        await YahooQuotes.GetAsync(symbols, Histories.None, "", ct).ConfigureAwait(false);
}
