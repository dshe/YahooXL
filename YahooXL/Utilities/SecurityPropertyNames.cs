using System.Reflection;

namespace YahooXL;

public class SecurityPropertyNames
{
    public static List<string> Get()
    {
        List<string> list = typeof(Security)
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Select(p => p.Name)
            .Where(n => n != "Properties")
            .Where(n => n != "DividendHistory" && n != "PriceHistory" && n != "PriceHistoryBase" && n != "SplitHistory")
            .Where(n => n != "ExchangeTimezone")
            .OrderBy(n => n)
            .ToList();

        if (list.Distinct(StringComparer.OrdinalIgnoreCase).Count() != list.Count)
            throw new InvalidOperationException("PropertyNames distinct count!");

        return list;
    }
}
