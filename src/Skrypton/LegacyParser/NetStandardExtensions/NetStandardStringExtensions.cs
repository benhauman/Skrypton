#if NETSTANDARD2_0
namespace System
{
    internal static class NetStandardStringExtensions
    {
        extension(string that)
        {
            internal string Replace(string oldValue, string newValue, StringComparison comparison)
            {
                return that.Replace(oldValue, newValue);
            }

            internal bool Contains(char c, StringComparison comparison)
            {
                return that.Contains(c.ToString());
            }

            internal int IndexOf(char c, StringComparison comparison)
            {
                return that.IndexOf(c);
            }

            internal bool StartsWith(char c, StringComparison comparison)
            {
                return that.StartsWith(c.ToString(), comparison);
            }

            internal bool EndsWith(char c, StringComparison comparison)
            {
                return that.EndsWith(c.ToString(), comparison);
            }

            internal string AsSpanX(int start) // ReadOnlySpan<char>
            {
                return that.Substring(start); //that.ToCharArray().AsSpan(start);
            }

            internal string AsSpanX(int start, int length) // ReadOnlySpan<char>
            {
                return that.Substring(start, length);// that.ToCharArray().AsSpan(start, length);
            }

            internal int GetHashCode(StringComparison comparison)
            {
                return that.GetHashCode();
            }
        }
    }
}
#endif