using System.Collections.Generic;
using System.Linq;

namespace Skrypton;

internal static class MyEnumerableExtensions
{
    internal static IReadOnlyCollection<T> ConcatCollection<T>(this IReadOnlyCollection<T> first, IReadOnlyCollection<T> second)
    {
        if (first.Count == 0)
            return second;
        if (second.Count == 0)
            return first;
        return first.Concat(second).ToArray();
    }
}