namespace YahooXL;

internal static class YahooFinance
{
    private static YahooQuotes YahooQuotes { get; }
    private static List<string> PropertyNamesList { get; }
    private static HashSet<string> PropertyNamesHash { get; }

    static YahooFinance()
    {
        YahooQuotes = new YahooQuotesBuilder()
                .WithLogger(AddIn.LogFactory.CreateLogger("YahooQuotes"))
                .Build();

        PropertyNamesList = typeof(Snapshot)
            .GetProperties()
            .OrderBy(pi => pi.MetadataToken)
            .Select(pi => pi.Name)
            .Where(name => name != "Properties")
            .ToList();
        
        PropertyNamesHash = new HashSet<string>(PropertyNamesList, StringComparer.OrdinalIgnoreCase);
    }

    internal static async Task<Dictionary<Symbol, Snapshot?>> GetSnapshotAsync(IEnumerable<Symbol> symbols, CancellationToken ct) =>
        await YahooQuotes.GetSnapshotAsync(symbols, ct).ConfigureAwait(false);

    internal static object GetValue(this Snapshot snapshot, string propertyName)
    {
        if (!snapshot.Properties.TryGetValue(propertyName, out object? value))
            return "";

        return value switch
        {
            Symbol symbol => symbol.Name,
            Instant instant => instant.ToDateTimeUtc().ToOADate(),
            _ => value ?? ""
        };
    }

    internal static string GetPropertyOrError(string propertyName)
    {
        if (PropertyNamesHash.TryGetValue(propertyName, out string? result))
            return result;
        if (int.TryParse(propertyName, out int i))
        {
            if (i >= 0 && i < PropertyNamesList.Count)
                return PropertyNamesList[i];
            return $"#Property name index: 0..{PropertyNamesList.Count - 1}.";
        }
        return $"#Invalid property name: '{propertyName}'.";
    }
}
