
using System;
using System.Linq;
using System.Collections;
using System.Reflection;

namespace Skrypton.Tests
{
    /// <summary>
    /// Default implementation of <see cref="T:System.Collections.Generic.IEqualityComparer`1" /> used by the xUnit.net equality assertions.
    /// </summary>
    /// <typeparam name="T">The type that is being compared.</typeparam>
    internal class AssertEqualityComparer<T> : System.Collections.Generic.IEqualityComparer<T>
    {
        private class TypeErasedEqualityComparer : System.Collections.IEqualityComparer
        {
            private readonly System.Collections.IEqualityComparer innerComparer;

            private static MethodInfo s_equalsMethod;

            public TypeErasedEqualityComparer(System.Collections.IEqualityComparer innerComparer)
            {
                this.innerComparer = innerComparer;
            }

            public new bool Equals(object x, object y)
            {
                if (x == null)
                {
                    return y == null;
                }
                if (y == null)
                {
                    return false;
                }
                Type type = (x.GetType() == y.GetType()) ? x.GetType() : typeof(object);
                if (AssertEqualityComparer<T>.TypeErasedEqualityComparer.s_equalsMethod == null)
                {
                    AssertEqualityComparer<T>.TypeErasedEqualityComparer.s_equalsMethod = typeof(AssertEqualityComparer<T>.TypeErasedEqualityComparer).GetTypeInfo().GetDeclaredMethod("EqualsGeneric");
                }
                return (bool)AssertEqualityComparer<T>.TypeErasedEqualityComparer.s_equalsMethod.MakeGenericMethod(new Type[]
                {
                    type
                }).Invoke(this, new object[]
                {
                    x,
                    y
                });
            }

            private bool EqualsGeneric<U>(U x, U y)
            {
                return new AssertEqualityComparer<U>(this.innerComparer).Equals(x, y);
            }

            public int GetHashCode(object obj)
            {
                throw new NotImplementedException();
            }
        }
        /*
        [CompilerGenerated]
        [Serializable]
        private sealed class xx_c
        {
            public static readonly AssertEqualityComparer<T>.xx_c<>9 = new AssertEqualityComparer<T>.xx_c();

            public static Func<Type, TypeInfo> <>9__10_0;

			public static Func<TypeInfo, bool> <>9__10_1;

			public static Func<TypeInfo, Type> <>9__10_2;

			internal TypeInfo<IsSet> b__10_0(Type i)
            {
                return i.GetTypeInfo();
            }

            internal bool <IsSet>b__10_1(TypeInfo ti)
            {
                return ti.IsGenericType;
            }

            internal Type<IsSet> b__10_2(TypeInfo ti)
            {
                return ti.GetGenericTypeDefinition();
            }
        }
        */
        private static readonly System.Collections.IEqualityComparer DefaultInnerComparer = new MyAssertEqualityComparerAdapter<object>(new AssertEqualityComparer<object>(null));

        private static readonly TypeInfo NullableTypeInfo = typeof(Nullable<>).GetTypeInfo();

        private readonly Func<System.Collections.IEqualityComparer> innerComparerFactory;

        private static MethodInfo s_compareTypedSetsMethod;

        /// <summary>
        /// Initializes a new instance of the <see cref="T:Xunit.Sdk.AssertEqualityComparer`1" /> class.
        /// </summary>
        /// <param name="innerComparer">The inner comparer to be used when the compared objects are enumerable.</param>
        public AssertEqualityComparer(System.Collections.IEqualityComparer innerComparer = null)
        {
            this.innerComparerFactory = (() => innerComparer ?? AssertEqualityComparer<T>.DefaultInnerComparer);
        }

        /// <inheritdoc />
        public bool Equals(T x, T y)
        {
            TypeInfo typeInfo = typeof(T).GetTypeInfo();
            if (!typeInfo.IsValueType || (typeInfo.IsGenericType && typeInfo.GetGenericTypeDefinition().GetTypeInfo().IsAssignableFrom(AssertEqualityComparer<T>.NullableTypeInfo)))
            {
                if (object.Equals(x, default(T)))
                {
                    return object.Equals(y, default(T));
                }
                if (object.Equals(y, default(T)))
                {
                    return false;
                }
            }
            IEquatable<T> equatable = x as IEquatable<T>;
            if (equatable != null)
            {
                return equatable.Equals(y);
            }
            IComparable<T> comparable = x as IComparable<T>;
            bool result;
            if (comparable != null)
            {
                try
                {
                    result = (comparable.CompareTo(y) == 0);
                    return result;
                }
                catch
                {
                }
            }
            IComparable comparable2 = x as IComparable;
            if (comparable2 != null)
            {
                try
                {
                    result = (comparable2.CompareTo(y) == 0);
                    return result;
                }
                catch
                {
                }
            }
            bool? flag = this.CheckIfDictionariesAreEqual(x, y);
            if (flag.HasValue)
            {
                return flag.GetValueOrDefault();
            }
            bool? flag2 = this.CheckIfSetsAreEqual(x, y, typeInfo);
            if (flag2.HasValue)
            {
                return flag2.GetValueOrDefault();
            }
            bool? flag3 = this.CheckIfEnumerablesAreEqual(x, y);
            if (flag3.HasValue)
            {
                if (!flag3.GetValueOrDefault())
                {
                    return false;
                }
                Array array = x as Array;
                Array array2 = y as Array;
                if (array != null && array2 != null)
                {
                    if (array.Rank != array2.Rank)
                    {
                        return false;
                    }
                    for (int i = 0; i < array.Rank; i++)
                    {
                        if (array.GetLength(i) != array2.GetLength(i))
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            else
            {
                IStructuralEquatable structuralEquatable = x as IStructuralEquatable;
                if (structuralEquatable != null && structuralEquatable.Equals(y, new AssertEqualityComparer<T>.TypeErasedEqualityComparer(this.innerComparerFactory())))
                {
                    return true;
                }
                TypeInfo typeInfo2 = typeof(IEquatable<>).MakeGenericType(new Type[]
                {
                    y.GetType()
                }).GetTypeInfo();
                if (typeInfo2.IsAssignableFrom(x.GetType().GetTypeInfo()))
                {
                    return (bool)typeInfo2.GetDeclaredMethod("Equals").Invoke(x, new object[]
                    {
                        y
                    });
                }
                TypeInfo typeInfo3 = typeof(IComparable<>).MakeGenericType(new Type[]
                {
                    y.GetType()
                }).GetTypeInfo();
                if (typeInfo3.IsAssignableFrom(x.GetType().GetTypeInfo()))
                {
                    MethodInfo declaredMethod = typeInfo3.GetDeclaredMethod("CompareTo");
                    try
                    {
                        result = ((int)declaredMethod.Invoke(x, new object[]
                        {
                            y
                        }) == 0);
                        return result;
                    }
                    catch
                    {
                    }
                }
                return object.Equals(x, y);
            }
            //return result;
        }

        private bool? CheckIfEnumerablesAreEqual(T x, T y)
        {
            IEnumerable enumerable = x as IEnumerable;
            IEnumerable enumerable2 = y as IEnumerable;
            bool? result;
            if (enumerable == null || enumerable2 == null)
            {
                result = null;
                return result;
            }
            IEnumerator enumerator = null;
            IEnumerator enumerator2 = null;
            try
            {
                enumerator = enumerable.GetEnumerator();
                enumerator2 = enumerable2.GetEnumerator();
                IEqualityComparer equalityComparer = this.innerComparerFactory();
                bool flag;
                bool flag2;
                while (true)
                {
                    flag = enumerator.MoveNext();
                    flag2 = enumerator2.MoveNext();
                    if (!flag || !flag2)
                    {
                        break;
                    }
                    if (!equalityComparer.Equals(enumerator.Current, enumerator2.Current))
                    {
                        goto Block_5;
                    }
                }
                result = new bool?(flag == flag2);
                return result;
            Block_5:
                result = new bool?(false);
            }
            finally
            {
                IDisposable disposable = enumerator as IDisposable;
                if (disposable != null)
                {
                    disposable.Dispose();
                }
                disposable = (enumerator2 as IDisposable);
                if (disposable != null)
                {
                    disposable.Dispose();
                }
            }
            return result;
        }

        private bool? CheckIfDictionariesAreEqual(T x, T y)
        {
            IDictionary dictionary = x as IDictionary;
            IDictionary dictionary2 = y as IDictionary;
            if (dictionary == null || dictionary2 == null)
            {
                bool? result = null;
                return result;
            }
            if (dictionary.Count != dictionary2.Count)
            {
                return new bool?(false);
            }
            IEqualityComparer equalityComparer = this.innerComparerFactory();
            System.Collections.Generic.HashSet<object> hashSet = new System.Collections.Generic.HashSet<object>(dictionary2.Keys.Cast<object>());
            foreach (object current in dictionary.Keys)
            {
                if (!hashSet.Contains(current))
                {
                    bool? result = new bool?(false);
                    return result;
                }
                object x2 = dictionary[current];
                object y2 = dictionary2[current];
                if (!equalityComparer.Equals(x2, y2))
                {
                    bool? result = new bool?(false);
                    return result;
                }
                hashSet.Remove(current);
            }
            return new bool?(hashSet.Count == 0);
        }

        private bool? CheckIfSetsAreEqual(T x, T y, TypeInfo typeInfo)
        {
            if (!this.IsSet(typeInfo))
            {
                return null;
            }
            IEnumerable enumerable = x as IEnumerable;
            IEnumerable enumerable2 = y as IEnumerable;
            if (enumerable == null || enumerable2 == null)
            {
                return null;
            }
            Type type;
            if (typeof(T).GenericTypeArguments.Length != 1)
            {
                type = typeof(object);
            }
            else
            {
                type = typeof(T).GenericTypeArguments[0];
            }
            if (AssertEqualityComparer<T>.s_compareTypedSetsMethod == null)
            {
                AssertEqualityComparer<T>.s_compareTypedSetsMethod = base.GetType().GetTypeInfo().GetDeclaredMethod("CompareTypedSets");
            }
            return new bool?((bool)AssertEqualityComparer<T>.s_compareTypedSetsMethod.MakeGenericMethod(new Type[]
            {
                type
            }).Invoke(this, new object[]
            {
                enumerable,
                enumerable2
            }));
        }

        private bool CompareTypedSets<R>(System.Collections.IEnumerable enumX, System.Collections.IEnumerable enumY)
        {
            System.Collections.Generic.HashSet<R> arg_18_0 = new System.Collections.Generic.HashSet<R>(enumX.Cast<R>());
            System.Collections.Generic.HashSet<R> equals = new System.Collections.Generic.HashSet<R>(enumY.Cast<R>());
            return arg_18_0.SetEquals(equals);
        }
        private bool IsSet(TypeInfo typeInfo)
        {
            //System.Collections.Generic.IEnumerable<Type> arg_25_0 = typeInfo.ImplementedInterfaces;
            ///Func<Type, TypeInfo> arg_25_1;
            return true;
            /*
                IEnumerable<Type> arg_25_0 = typeInfo.ImplementedInterfaces;
                Func<Type, TypeInfo> arg_25_1;
                if ((arg_25_1 = AssertEqualityComparer<T>.<> c.<> 9__10_0) == null)
                {
                    arg_25_1 = (AssertEqualityComparer<T>.<> c.<> 9__10_0 = new Func<Type, TypeInfo>(AssertEqualityComparer<T>.<> c.<> 9.< IsSet > b__10_0));
                }
                IEnumerable<TypeInfo> arg_49_0 = arg_25_0.Select(arg_25_1);
                Func<TypeInfo, bool> arg_49_1;
                if ((arg_49_1 = AssertEqualityComparer<T>.<> c.<> 9__10_1) == null)
                {
                    arg_49_1 = (AssertEqualityComparer<T>.<> c.<> 9__10_1 = new Func<TypeInfo, bool>(AssertEqualityComparer<T>.<> c.<> 9.< IsSet > b__10_1));
                }
                IEnumerable<TypeInfo> arg_6D_0 = arg_49_0.Where(arg_49_1);
                Func<TypeInfo, Type> arg_6D_1;
                if ((arg_6D_1 = AssertEqualityComparer<T>.<> c.<> 9__10_2) == null)
                {
                    arg_6D_1 = (AssertEqualityComparer<T>.<> c.<> 9__10_2 = new Func<TypeInfo, Type>(AssertEqualityComparer<T>.<> c.<> 9.< IsSet > b__10_2));
                }
                return arg_6D_0.Select(arg_6D_1).Contains(typeof(ISet<>).GetGenericTypeDefinition());
            */
        }
        /// <inheritdoc />
        public int GetHashCode(T obj)
        {
            throw new NotImplementedException();
        }
    }
}
