using System.Linq;
using YahooQuotesApi;

#nullable enable

namespace YahooXL
{
    internal static class Extensions
    {
        internal static string GetFieldNameOrNotFound(Security security, string fieldName) =>
           int.TryParse(fieldName, out int i) ? GetFieldNameFromIndex(security, i) : $"Field not found: \"{fieldName}\".";

        private static string GetFieldNameFromIndex(Security security, int index) => // slow!
            security.Fields.Keys.OrderBy(k => k).ElementAtOrDefault(index) ?? $"Invalid field index: {index}.";
    }

}
