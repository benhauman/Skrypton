using System;

namespace Skrypton.Tests
{
    /// <summary>
	/// A class that wraps <see cref="T:System.Collections.Generic.IEqualityComparer`1" /> to create <see cref="T:System.Collections.IEqualityComparer" />.
	/// </summary>
	/// <typeparam name="T">The type that is being compared.</typeparam>
	internal class MyAssertEqualityComparerAdapter<T> : System.Collections.IEqualityComparer
    {
        private readonly System.Collections.Generic.IEqualityComparer<T> innerComparer;

        /// <summary>
        /// Initializes a new instance of the <see cref="T:Xunit.Sdk.AssertEqualityComparerAdapter`1" /> class.
        /// </summary>
        /// <param name="innerComparer">The comparer that is being adapted.</param>
        public MyAssertEqualityComparerAdapter(System.Collections.Generic.IEqualityComparer<T> innerComparer)
        {
            this.innerComparer = innerComparer;
        }

        /// <inheritdoc />
        public new bool Equals(object x, object y)
        {
            return this.innerComparer.Equals((T)((object)x), (T)((object)y));
        }

        /// <inheritdoc />
        public int GetHashCode(object obj)
        {
            throw new NotImplementedException();
        }
    }
}
