using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport
{
    public static class IAccessValuesUsingVBScriptRulesExtensions
    {
        internal const int MaxNumberOfMemberAccessorBeforeArraysRequired = 5;

        internal static void SETm1argp(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target, string memberAccessor, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));
            if (string.IsNullOrEmpty(memberAccessor)) throw new ArgumentException("Value cannot be null or empty.", nameof(memberAccessor));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: memberAccessor, argumentProviderBuilder.GetArgs());
        }

        // This one allows for the arguments to not be mentioned at all if they're not required for a SET (unlike CALL, there is no concept of "forced brackets"
        // when there are zero arguments since a SET is always part of a value-setting statement, which means that the brackets are an essential part of the
        // statement and not optional tokens that may or may not be present)
        public static void SETm1a0(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target, string memberAccessor)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(memberAccessor)) throw new ArgumentException("Value cannot be null or empty.", nameof(memberAccessor));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: memberAccessor, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty);
        }
        public static void SETm0a0(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: null, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty);
        }
        public static void SETm1a1(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target, string memberAccessor, object arg1)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(memberAccessor)) throw new ArgumentException("Value cannot be null or empty.", nameof(memberAccessor));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: memberAccessor, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [arg1]));
        }
        public static void SETm0a1(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target, object arg1)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: null, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [arg1]));
        }
        // This one allows for no arguments OR member accessors to be mentioned - this should only be used for errors cases (since, otherwise, a simple assignment
        // would be more appropriate, no SET call would be required at all). This may be used for the representation of "a = 1" where "a" is a function or a
        // constant, the translated output would be call to this function where the target to would actually be a call to RAISEERROR so that the valueToSet
        // may be evaluated and then a can-not-set-this error raised (consistent with how VBScript would handle it).
        internal static void SETnm(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object? context, object target)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            source.SET(valueToSetTo, context, target, optionalMemberAccessor: null, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty);
        }

        // Convenience methods for when there are no arguments (supporting up to MaxNumberOfMemberAccessorBeforeArraysRequired members accessors, just as the
        // extension methods further down do which look after the with-arguments signatures)
        public static object? CALLm0v0(this IAccessValuesUsingVBScriptRules source, object context, object target)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, [], ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: 0);
        }
        public static object? CALLm1v0(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, [CallerLineNumber] int line = 0)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: line);
        }
        public static object? CALLm1v1(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, object value1)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1]), line: 0);
        }
        public static object? CALLm1v2(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, object value1, object value2)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2]), line: 0);
        }
        public static object? CALLm1v3(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, object value1, object value2, object value3)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3]), line: 0);
        }
        public static object? CALLm1v4(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, object value1, object value2, object value3, object value4)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3, value4]), line: 0);
        }
        public static object? CALLm1v5(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, object value1, object value2, object value3, object value4, object value5)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3, value4, value5]), line: 0);
        }
        public static object? CALLm2v0(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1, member2 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: 0);
        }
        public static object? CALLm2v1(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, object value1)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1, member2 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1]), line: 0);
        }
        //public static object? CALLm2v2(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, object value1, object value2)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2]), line: 0);
        //}
        //public static object? CALLm2v3(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, object value1, object value2, object value3)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3]), line: 0);
        //}
        //public static object? CALLm2v4(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, object value1, object value2, object value3, object value4)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3, value4]), line: 0);
        //}
        //public static object? CALLm2v5(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, object value1, object value2, object value3, object value4, object value5)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1, value2, value3, value4, value5]), line: 0);
        //}
        //public static object? CALLm3v0(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2, member3 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: 0);
        //}
        //public static object? CALLm4v0(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2, member3, member4 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: 0);
        //}
        //public static object? CALLm5v0(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, string member5)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //
        //    return source.CALL(context, target, new[] { member1, member2, member3, member4, member5 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBracketsEmpty, line: 0);
        //}

        public static object? CALLm3v1(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, object value1)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            return source.CALL(context, target, new[] { member1, member2, member3 }, DefaultCallArgumentProvider.CreateArgumentProviderForValues(useBracketsWhereZeroArguments: false, [value1]), line: 0);
        }

        // Convenience methods so that the calling code can omit the "GetArgs" call if an IBuildCallArgumentProviders is already available (results in shorter
        // translated code)
        public static object? CALLarrmargp(this IAccessValuesUsingVBScriptRules source, object context, object target, string[] members, IBuildCallArgumentProviders argumentProviderBuilder, [CallerLineNumber] int line = 0)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));

            return source.CALL(context, target, members, argumentProviderBuilder.GetArgs(), line: line);
        }
        // Convenience methods for when there are a known number of accessor members (including zero) and arguments - providing the argument builder means that
        // the translated code can be shorter (since there will be less "GetArgs" calls) but the trust is placed in these extension methods that the arguments
        // set will not be manipulated (extended). Since there would already trust that these won't manipulate any values if IProvideCallArguments references
        // were passed then this isn't a big deal (strictly speaking these methods express requirements greater than they really need but the shorter code is
        // worth it).
        public static object? CALLm0argp(this IAccessValuesUsingVBScriptRules source, object context, object target, IBuildCallArgumentProviders argumentProviderBuilder, [CallerLineNumber] int line = 0)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));

            return source.CALL(context, target, [], argumentProviderBuilder.GetArgs(), line);
        }
        public static object? CALLm1argp(this IAccessValuesUsingVBScriptRules source, object? context, object target, string member1, IBuildCallArgumentProviders argumentProviderBuilder)//, [CallerLineNumber] int line = 0)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));

            return source.CALL(context, target, new[] { member1 }, argumentProviderBuilder.GetArgs(), 0);
        }
        public static object? CALLm2argp(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, IBuildCallArgumentProviders argumentProviderBuilder, [CallerLineNumber] int line = 0)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));

            return source.CALL(context, target, new[] { member1, member2 }, argumentProviderBuilder.GetArgs(), line);
        }
        public static object? CALLm3argp(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilder));

            return source.CALL(context, target, new[] { member1, member2, member3 }, argumentProviderBuilder.GetArgs(), line: 0);
        }
        //public static object? CALLm4argp(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, IBuildCallArgumentProviders argumentProviderBuilder)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //    if (argumentProviderBuilder == null)
        //        throw new ArgumentNullException(nameof(argumentProviderBuilder));
        //
        //    return source.CALL(context, target, new[] { member1, member2, member3, member4 }, argumentProviderBuilder.GetArgs(), line: 0);
        //}
        //public static object? CALLm5argp(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, string member5, IBuildCallArgumentProviders argumentProviderBuilder)
        //{
        //    if (source == null)
        //        throw new ArgumentNullException(nameof(source));
        //    if (argumentProviderBuilder == null)
        //        throw new ArgumentNullException(nameof(argumentProviderBuilder));
        //
        //    return source.CALL(context, target, new[] { member1, member2, member3, member4, member5 }, argumentProviderBuilder.GetArgs(), line: 0);
        //}

        private sealed class ZeroArgumentArgumentProvider : IProvideCallArguments
        {
            //internal static IProvideCallArguments WithEnforcedArgumentBracketsEmpty = new ZeroArgumentArgumentProvider(true);
            internal static IProvideCallArguments WithoutEnforcedArgumentBracketsEmpty = new ZeroArgumentArgumentProvider(false);
            private ZeroArgumentArgumentProvider(bool useBracketsWhereZeroArguments)
            {
                UseBracketsWhereZeroArguments = useBracketsWhereZeroArguments;
            }

            public int NumberOfArguments { get { return 0; } }

            public bool UseBracketsWhereZeroArguments { get; private set; }

            public object[] GetInitialValues()
            {
                return [];
            }

            public void OverwriteValueIfByRef(int index, object value)
            {
                throw new ArgumentException("There are no arguments to overwrite");
            }
        }
    }
}
