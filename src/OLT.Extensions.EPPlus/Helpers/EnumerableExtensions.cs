using System.Collections.Generic;
using System.Diagnostics;

namespace OLT.Extensions.EPPlus.Helpers
{
    internal static class EnumerableExtensions
    {
        [DebuggerStepThrough]
        internal static bool IsGreaterThanOne<T>(this IEnumerable<T> source)
        {
            using (IEnumerator<T> enumerator = source.GetEnumerator())
            {
                if (enumerator.MoveNext())
                {
                    return enumerator.MoveNext();
                }
                return false;
            }
        }
    }
}