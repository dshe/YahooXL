using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YahooFinanceApi;

#nullable enable

namespace YahooDelayedQuotesXLAddIn
{
    internal static class Extensions
    {
        internal static string GetFieldNameOrNotFound(Security security, string fieldName) =>
           int.TryParse(fieldName, out int i) ? GetFieldNameFromIndex(security, i) : $"Field not found: \"{fieldName}\".";

        private static string GetFieldNameFromIndex(Security security, int index) => // slow!
            security.Fields.Keys.OrderBy(k => k).ElementAtOrDefault(index) ?? $"Invalid field index: {index}.";

    }
}
