using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// How selectNodes / selectSingleNode evaluate a query (issue #94), set with DOMDocument.setProperty: MSXML 3's default selection language
    /// "XSLPattern" matches names as they are written ("dt:CompanyName", "plain" also in a default namespace) and ignores namespace URIs; "XPath"
    /// resolves the prefixes through the "SelectionNamespaces" property ("xmlns:q='urn:x'") and raises an error for an undeclared prefix
    /// </summary>
    internal sealed class MsxmlSelection
    {
        internal const string ErrorSource = "msxml3.dll";
        internal const int ErrorNumber = unchecked((int)0x80004005); // E_FAIL, as MSXML reports these errors to VBScript
        public const string XPath = "XPath";
        public const string XslPattern = "XSLPattern";

        private static readonly ConditionalWeakTable<XmlDocument, MsxmlSelection> Selections = new ConditionalWeakTable<XmlDocument, MsxmlSelection>();
        private static readonly Regex NamespaceDeclaration = new Regex(@"\G\s*xmlns(?::(?<prefix>[A-Za-z_][\w.\-]*))?\s*=\s*(?:'(?<uri>[^']*)'|""(?<uri>[^""]*)"")\s*", RegexOptions.CultureInvariant);

        private Dictionary<string, string> _namespaces = new Dictionary<string, string>(StringComparer.Ordinal);

        public string Language { get; private set; } = XslPattern;

        public string Namespaces { get; private set; } = "";

        public static MsxmlSelection For(XmlDocument document) => Selections.GetValue(document, _ => new MsxmlSelection());

        private static MsxmlSelection For(XmlNode context) => For((context as XmlDocument) ?? context.OwnerDocument!);

        public static CustomException Error(string description) => new CustomException(ErrorNumber, ErrorSource, description + "\r\n"); // MSXML ends its descriptions with a line break

        public void SetLanguage(object value)
        {
            // the values are case-sensitive, as with MSXML
            if (!(value is string language) || ((language != XPath) && (language != XslPattern)))
                throw Error("Property value is invalid.");
            Language = language;
        }

        /// <summary>
        /// "xmlns:q='urn:x' xmlns:r='urn:y'" - the prefixes that XPath queries may use
        /// </summary>
        public void SetNamespaces(object value)
        {
            string declarations = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            var namespaces = new Dictionary<string, string>(StringComparer.Ordinal);
            int position = 0;
            while (position < declarations.Length)
            {
                Match match = NamespaceDeclaration.Match(declarations, position);
                if (!match.Success || (match.Length == 0))
                    throw Error("A name contained an invalid character.");
                if (match.Groups["prefix"].Success)
                    namespaces[match.Groups["prefix"].Value] = match.Groups["uri"].Value;
                position += match.Length;
            }
            _namespaces = namespaces;
            Namespaces = declarations;
        }

        public static XmlNodeList? SelectNodes(XmlNode context, string query)
        {
            if (string.IsNullOrEmpty(query))
                return null;
            MsxmlSelection selection = For(context);
            return (selection.Language == XslPattern)
                ? context.SelectNodes(ToNameMatching(query))
                : context.SelectNodes(query, selection.NamespaceManager(context, query));
        }

        public static XmlNode? SelectSingleNode(XmlNode context, string query)
        {
            if (string.IsNullOrEmpty(query))
                return null;
            MsxmlSelection selection = For(context);
            return (selection.Language == XslPattern)
                ? context.SelectSingleNode(ToNameMatching(query))
                : context.SelectSingleNode(query, selection.NamespaceManager(context, query));
        }

        private XmlNamespaceManager NamespaceManager(XmlNode context, string query)
        {
            XmlNameTable names = ((context as XmlDocument) ?? context.OwnerDocument!).NameTable;
            var manager = new XmlNamespaceManager(names);
            foreach (KeyValuePair<string, string> declaration in _namespaces)
                manager.AddNamespace(declaration.Key, declaration.Value);
            foreach (Token token in Tokenize(query))
            {
                int separator = token.Text.IndexOf(':');
                if ((token.Kind == TokenKind.NameTest) && (separator > 0) && !_namespaces.ContainsKey(token.Text.Substring(0, separator)))
                    throw Error($"Reference to undeclared namespace prefix: '{token.Text.Substring(0, separator)}'.");
            }
            return manager;
        }

        /// <summary>
        /// The XPath query with each name test matching the name as it is written (XSLPattern): "dt:c" is *[name()='dt:c'], "dt:*" is
        /// *[starts-with(name(),'dt:')] (also after "@" and an axis); an index is 0-based in XSLPattern ("item[0]" is the first item)
        /// </summary>
        internal static string ToNameMatching(string query)
        {
            var text = new StringBuilder(query.Length * 2);
            List<Token> tokens = new List<Token>(Tokenize(query));
            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                if (token.Kind != TokenKind.NameTest)
                {
                    text.Append(IsIndex(tokens, i, out int index) ? (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) : token.Text);
                    continue;
                }
                if (token.Text.EndsWith(":*", StringComparison.Ordinal))
                    text.Append("*[starts-with(name(),'").Append(token.Text, 0, token.Text.Length - 1).Append("')]");
                else
                    text.Append("*[name()='").Append(token.Text).Append("']");
            }
            return text.ToString();
        }

        /// <summary>
        /// Is the token a whole number that is the only content of a predicate ("[1]")
        /// </summary>
        private static bool IsIndex(List<Token> tokens, int position, out int index)
        {
            index = 0;
            if ((tokens[position].Kind != TokenKind.Operand) || !int.TryParse(tokens[position].Text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out index))
                return false;
            int before = position - 1;
            while ((before >= 0) && string.IsNullOrWhiteSpace(tokens[before].Text))
                before--;
            int after = position + 1;
            while ((after < tokens.Count) && string.IsNullOrWhiteSpace(tokens[after].Text))
                after++;
            return (before >= 0) && (tokens[before].Text == "[") && (after < tokens.Count) && (tokens[after].Text == "]");
        }

        private enum TokenKind
        {
            Other,
            Operand, // a literal, a number, ")", "]", ".", "..", "*" (any name), a variable
            NameTest, // a (qualified) name that selects nodes
        }

        private readonly struct Token
        {
            public Token(TokenKind kind, string text)
            {
                Kind = kind;
                Text = text;
            }

            public TokenKind Kind { get; }

            public string Text { get; }
        }

        /// <summary>
        /// The tokens of an XPath query, as the XPath 1.0 lexer distinguishes them: a name is an operator ("and", "div", ..) after an operand,
        /// a function or node type before "(", an axis before "::" and a name test otherwise
        /// </summary>
        private static IEnumerable<Token> Tokenize(string query)
        {
            bool afterOperand = false;
            int i = 0;
            while (i < query.Length)
            {
                char c = query[i];
                int start = i;
                if (char.IsWhiteSpace(c))
                {
                    while ((i < query.Length) && char.IsWhiteSpace(query[i]))
                        i++;
                    yield return new Token(TokenKind.Other, query.Substring(start, i - start));
                    continue;
                }
                if ((c == '\'') || (c == '"'))
                {
                    int end = query.IndexOf(c, i + 1);
                    i = (end < 0) ? query.Length : end + 1;
                    afterOperand = true;
                    yield return new Token(TokenKind.Operand, query.Substring(start, i - start));
                    continue;
                }
                if (char.IsDigit(c) || ((c == '.') && (i + 1 < query.Length) && char.IsDigit(query[i + 1])))
                {
                    while ((i < query.Length) && (char.IsDigit(query[i]) || (query[i] == '.')))
                        i++;
                    afterOperand = true;
                    yield return new Token(TokenKind.Operand, query.Substring(start, i - start));
                    continue;
                }
                if (IsNameStart(c) || (c == '$'))
                {
                    i++;
                    while ((i < query.Length) && IsNameChar(query[i]))
                        i++;
                    if ((i + 1 < query.Length) && (query[i] == ':') && (query[i + 1] != ':') && (c != '$'))
                    {
                        // a prefix: "dt:c" or "dt:*"
                        i++;
                        if (query[i] == '*')
                            i++;
                        else
                            while ((i < query.Length) && IsNameChar(query[i]))
                                i++;
                    }
                    string name = query.Substring(start, i - start);
                    if (c == '$')
                    {
                        afterOperand = true;
                        yield return new Token(TokenKind.Operand, name);
                        continue;
                    }
                    if (afterOperand)
                    {
                        afterOperand = false; // and, or, div, mod
                        yield return new Token(TokenKind.Other, name);
                        continue;
                    }
                    int next = i;
                    while ((next < query.Length) && char.IsWhiteSpace(query[next]))
                        next++;
                    bool isFunction = (next < query.Length) && (query[next] == '(');
                    bool isAxis = (next + 1 < query.Length) && (query[next] == ':') && (query[next + 1] == ':');
                    afterOperand = !(isFunction || isAxis);
                    yield return new Token((isFunction || isAxis) ? TokenKind.Other : TokenKind.NameTest, name);
                    continue;
                }
                if (c == '*')
                {
                    i++;
                    TokenKind kind = afterOperand ? TokenKind.Other : TokenKind.Operand; // multiplication or any name
                    afterOperand = !afterOperand;
                    yield return new Token(kind, "*");
                    continue;
                }
                if (c == '.')
                {
                    i += ((i + 1 < query.Length) && (query[i + 1] == '.')) ? 2 : 1;
                    afterOperand = true;
                    yield return new Token(TokenKind.Operand, query.Substring(start, i - start));
                    continue;
                }
                i++;
                if ((i < query.Length) && (((c == ':') && (query[i] == ':')) || ((c == '/') && (query[i] == '/')) || (((c == '!') || (c == '<') || (c == '>')) && (query[i] == '='))))
                    i++;
                afterOperand = (c == ')') || (c == ']');
                yield return new Token(TokenKind.Other, query.Substring(start, i - start));
            }
        }

        private static bool IsNameStart(char c) => char.IsLetter(c) || (c == '_');

        private static bool IsNameChar(char c) => char.IsLetterOrDigit(c) || (c == '_') || (c == '-') || (c == '.');
    }
}
