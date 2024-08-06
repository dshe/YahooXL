using System.Collections.Immutable;
using System.Reactive.Linq;

namespace YahooXL;

public static class Securities
{
    private static readonly string[] ExcludedProperties = ["Props", "DividendHistory", "PriceHistory", "PriceHistoryBase", "SplitHistory"];

    private static readonly Lazy<List<string>> PropertyNames = new(() =>
        typeof(Security)
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .OrderBy(p => p.MetadataToken)
            .Select(p => p.Name)
            .Where(n => !ExcludedProperties.Contains(n, StringComparer.OrdinalIgnoreCase))
            .ToList());

    internal static bool PropertyExists(string name) => PropertyNames.Value.Contains(name, StringComparer.OrdinalIgnoreCase);

    internal static string? GetPropertyNameByIndex(int i, Security? security = null)
    {
        if (i < 0)
            return null;
        int count = PropertyNames.Value.Count;
        if (i < count)
            return PropertyNames.Value[i];
        if (security == null) 
            return null;
        return security.Props.Values
            .Where(x => x.Category == PropCategory.New)
            .Select(x => x.Name)
            .OrderBy(x => x)
            .ElementAtOrDefault(i - count);
    }

    internal static object GetValue(Security security, string propertyName)
    {
        if (int.TryParse(propertyName, out int i))
        {
            string? name = GetPropertyNameByIndex(i, security);
            if (name == null)
                return "YahooXL: Invalid propery index.";
            return name;
        }

        if (!security.Props.TryGetValue(propertyName, out Prop? p))
            return "YahooXL: Unknown propery.";

        // PropCategory: Expected, Calculated, Missing, New
        object? v = p.Value;
        if (v is Symbol symbol)
            return symbol.Name;
        return v ?? "";
    }
}
