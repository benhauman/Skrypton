using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.Tests.Shared.Comparers
{
    public class TokenComparer : IEqualityComparer<IToken>
    {
        internal static readonly TokenComparer Instance = new TokenComparer();
        public bool Equals(IToken x, IToken y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.NumericValueToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.NumericValueToken.CompareNumericValueToken(
                      (Skrypton.LegacyParser.Tokens.Basic.NumericValueToken)x,
                      (Skrypton.LegacyParser.Tokens.Basic.NumericValueToken)y) == 0;
            }
            //if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.NameToken))
            if (x is Skrypton.LegacyParser.Tokens.Basic.NameToken) // 'DoNotRenameNameToken'
            {
                return Skrypton.LegacyParser.Tokens.Basic.NameToken.CompareNameTokens(
                      (Skrypton.LegacyParser.Tokens.Basic.NameToken)x,
                      (Skrypton.LegacyParser.Tokens.Basic.NameToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.OpenBrace))
            {
                return Skrypton.LegacyParser.Tokens.Basic.OpenBrace.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.OpenBrace)x,
                    (Skrypton.LegacyParser.Tokens.Basic.OpenBrace)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.CloseBrace))
            {
                return Skrypton.LegacyParser.Tokens.Basic.CloseBrace.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.CloseBrace)x,
                    (Skrypton.LegacyParser.Tokens.Basic.CloseBrace)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.OperatorToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.OperatorToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.OperatorToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.OperatorToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.ArgumentSeparatorToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.ArgumentSeparatorToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.ArgumentSeparatorToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.ArgumentSeparatorToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.LogicalOperatorToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.LogicalOperatorToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.LogicalOperatorToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.LogicalOperatorToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.ComparisonOperatorToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.ComparisonOperatorToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.ComparisonOperatorToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.ComparisonOperatorToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.BuiltInFunctionToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.BuiltInFunctionToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.BuiltInFunctionToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.BuiltInFunctionToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.MemberAccessorOrDecimalPointToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.BuiltInFunctionToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.MemberAccessorOrDecimalPointToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.MemberAccessorOrDecimalPointToken)y) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.StringToken))
            {
                return string.CompareOrdinal(
                    ((Skrypton.LegacyParser.Tokens.Basic.StringToken)x).Content,
                    ((Skrypton.LegacyParser.Tokens.Basic.StringToken)x).Content) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.BuiltInValueToken))
            {
                return Skrypton.LegacyParser.Tokens.Basic.BuiltInValueToken.CompareAtomTokens(
                    (Skrypton.LegacyParser.Tokens.Basic.BuiltInValueToken)x,
                    (Skrypton.LegacyParser.Tokens.Basic.BuiltInValueToken)y) == 0;
            }
            if (x is Skrypton.LegacyParser.Tokens.Basic.CommentToken)
            {
                return string.CompareOrdinal(
                    ((Skrypton.LegacyParser.Tokens.Basic.CommentToken)x).Content,
                    ((Skrypton.LegacyParser.Tokens.Basic.CommentToken)x).Content) == 0;
            }
            if (x.GetType() == typeof(Skrypton.LegacyParser.Tokens.Basic.UnprocessedContentToken))
            {
                return string.CompareOrdinal(
                    ((Skrypton.LegacyParser.Tokens.Basic.UnprocessedContentToken)x).Content,
                    ((Skrypton.LegacyParser.Tokens.Basic.UnprocessedContentToken)x).Content) == 0;
            }
            if (x is Skrypton.LegacyParser.Tokens.Basic.AbstractEndOfStatementToken)
            {
                return string.CompareOrdinal(
                    ((Skrypton.LegacyParser.Tokens.Basic.AbstractEndOfStatementToken)x).Content,
                    ((Skrypton.LegacyParser.Tokens.Basic.AbstractEndOfStatementToken)x).Content) == 0;
            }
            if (x is Skrypton.LegacyParser.Tokens.Basic.KeyWordToken)
            {
                return string.CompareOrdinal(
                    ((Skrypton.LegacyParser.Tokens.Basic.KeyWordToken)x).Content,
                    ((Skrypton.LegacyParser.Tokens.Basic.KeyWordToken)x).Content) == 0;
            }
            if (x is Skrypton.StageTwoParser.Tokens.MemberAccessorToken)
            {
                return string.CompareOrdinal(
                    ((Skrypton.StageTwoParser.Tokens.MemberAccessorToken)x).Content,
                    ((Skrypton.StageTwoParser.Tokens.MemberAccessorToken)x).Content) == 0;
            }
            IComparable comparableX = (IComparable)x;
            IComparable comparableY = (IComparable)y;
            var cmp = comparableX.CompareTo(comparableY);
            if (cmp == 0)
                return true;
            return false;

            //lubo:            var bytesX = GetBytes(x);
            //lubo:            var bytesY = GetBytes(y);
            //lubo:            if (bytesX.Length != bytesY.Length)
            //lubo:                return false;
            //lubo:            for (var indexBytes = 0; indexBytes < bytesX.Length; indexBytes++)
            //lubo:            {
            //lubo:                if (bytesX[indexBytes] != bytesY[indexBytes])
            //lubo:                    return false;
            //lubo:            }
            //lubo:            return true;
            throw new NotImplementedException($"x:{x.GetType().Name}, y:{y.GetType().Name}");
        }

        //lubo:private static byte[] GetBytes(object obj)
        //lubo:{
        //lubo:    if (obj == null)
        //lubo:        throw new ArgumentNullException("obj");
        //lubo:
        //lubo:    using (var stream = new MemoryStream())
        //lubo:    {
        //lubo:        var formatter = new BinaryFormatter();
        //lubo:        formatter.Serialize(stream, obj);
        //lubo:        stream.Seek(0, SeekOrigin.Begin);
        //lubo:        return ReadBytesFromStream(stream);
        //lubo:    }
        //lubo:}

        private static byte[] ReadBytesFromStream(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            var buffer = new byte[4096];
            var read = 0;
            int chunk;
            while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                read += chunk;
                if (read == buffer.Length)
                {
                    int nextByte = stream.ReadByte();
                    if (nextByte == -1)
                        return buffer;
                    byte[] newBuffer = new byte[buffer.Length * 2];
                    Array.Copy(buffer, newBuffer, buffer.Length);
                    newBuffer[read] = (byte)nextByte;
                    buffer = newBuffer;
                    read++;
                }
            }
            var ret = new byte[read];
            Array.Copy(buffer, ret, read);
            return ret;
        }

        public int GetHashCode(IToken obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
