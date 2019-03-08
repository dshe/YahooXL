using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#nullable enable

namespace YahooXL.Utilities
{
    static class Utility
    {
        internal static List<string> CaseInsensitiveDuplicates(this IEnumerable<string> strings)
        {
            var hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return strings.Where(str => !hashSet.Add(str)).ToList();
        }

    }
}
