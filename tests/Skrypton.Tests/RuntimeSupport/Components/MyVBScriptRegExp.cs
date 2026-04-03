using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security.Authentication;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [SourceClassName("RegExp")] // for TYPENAME(CreateObject("VBScript.RegExp"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]

    internal sealed class MyVBScriptRegExp : IRegExp2
    {
        // =============================
        //   Fields stored VBScript-style
        // =============================
        private string _pattern = "";
        private bool _ignoreCase = false;
        private bool _global = false;      // In VBScript RegExp, Replace only replaces the first match unless you set Global = true.
        private bool _multiline = false;   // VBScript uses MultiLine flag
        private int _compare = 0;          // 0 = Binary; 1 = Text (culture-aware)

        private readonly CultureInfo _culture;
        internal MyVBScriptRegExp(CultureInfo culture)
        {
            _culture = culture;
        }

        // =============================
        //      Properties
        // =============================

        public string Pattern
        {
            get => _pattern;
            set => _pattern = value ?? "";
        }

        public bool IgnoreCase
        {
            get => _ignoreCase;
            set => _ignoreCase = value;
        }

        public bool Global
        {
            get => _global;
            set => _global = value;
        }

        public int Multiline
        {
            get => _multiline ? 1 : 0;
            set => _multiline = (value != 0);
        }

        public int Compare
        {
            get => _compare;
            set => _compare = value;
        }

        // =============================
        //     Core Method: Build Regex
        // =============================
        private Regex BuildRegex()
        {
            var opts = RegexOptions.None;

            if (_ignoreCase)
                opts |= RegexOptions.IgnoreCase;
            if (_multiline)
                opts |= RegexOptions.Multiline;

            // VBScript Compare: 0=Binary,1=Text
            if (_compare == 1)
                opts |= RegexOptions.CultureInvariant;

            return new Regex(_pattern, opts);
        }

        // =============================
        //     VBScript: Test()
        // =============================
        public bool Test(string input)
        {
            if (input == null)
                input = "";

            // .NET regex does not perform culture‑aware case folding.
            // Because of that, "straße" ⇄ "STRASSE" is not considered equal by Regex IgnoreCase.
            if (_compare == 0) // not culture invariant (0:binary compare / literal) ? (1:culture specific)
            {
                int cmp = _ignoreCase ? string.Compare(Pattern, input, _culture, CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreCase)
                                      : string.Compare(Pattern, input, _culture, CompareOptions.IgnoreNonSpace);
                if (cmp == 0)
                    return true;
            }


            return BuildRegex().IsMatch(input);
        }

        // =============================
        //   VBScript: Execute()
        // =============================
        public IMatchCollection Execute(string input)
        {
            if (input == null)
                input = "";

            Regex regex = BuildRegex();
            var matches = regex.Matches(input);

            return new MyMatchCollection(matches);
        }

        // =============================
        //   VBScript: Replace()
        // =============================
        public string Replace(string input, string replacement)
        {
            if (input == null)
                input = "";
            if (replacement == null)
                replacement = "";

            var regex = BuildRegex();

            if (_global)
                return regex.Replace(input, replacement);
            else
                return regex.Replace(input, replacement, 1);
        }
    }
    internal sealed class MyMatch : IMatch
    {
        private readonly Match _m;

        public MyMatch(Match m) => _m = m;

        public string Value => _m.Value;
        public int FirstIndex => _m.Index;
        public int Length => _m.Length;
        public ISubMatches SubMatches => new MySubMatches(_m);
    }

    internal sealed class MyMatchCollection : IMatchCollection
    {
        private readonly MatchCollection _matches;

        public MyMatchCollection(MatchCollection matches)
            => _matches = matches;

        public int Count => _matches.Count;

        public IMatch this[int index] => new MyMatch(_matches[index]);

        public System.Collections.IEnumerator GetEnumerator()
        {
            foreach (Match m in _matches)
                yield return new MyMatch(m);
        }
    }

    internal sealed class MySubMatches : ISubMatches
    {
        private readonly GroupCollection _groups;

        public MySubMatches(Match m) => _groups = m.Groups;

        public int Count => Math.Max(0, _groups.Count - 1);

        public object this[int index] => _groups[index + 1].Value;

        public System.Collections.IEnumerator GetEnumerator()
        {
            for (int i = 1; i < _groups.Count; i++)
                yield return _groups[i].Value;
        }
    }

    public enum RegExpFlags
    {
        // VBScript uses Boolean properties instead of a single flag enum.
        // This enum is left empty intentionally.
    }

    //[ComImport]
    //[Guid("3F4DACA3-160D-11D2-A8E9-00104B365C9F")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRegExp
    {
        // Properties
        string Pattern { get; set; }
        bool IgnoreCase { get; set; }
        bool Global { get; set; }

        // Methods
        bool Test(string input);
        IMatchCollection Execute(string input);
        string Replace(string input, string replacement);
    }

    //[ComImport]
    //[Guid("3F4DACA4-160D-11D2-A8E9-00104B365C9F")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRegExp2 : IRegExp
    {
        //// Basic settings
        //string Pattern { get; set; }
        //bool IgnoreCase { get; set; }
        //bool Global { get; set; }

        // New in IRegExp2
        int Multiline { get; set; }
        int Compare { get; set; }

        // Methods
        //bool Test(string input);
        //IMatchCollection Execute(string input);
        //string Replace(string input, string replacement);
    }
    //[ComImport]
    //[Guid("3F4DACA1-160D-11D2-A8E9-00104B365C9F")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMatch
    {
        string Value { get; }
        int FirstIndex { get; }
        int Length { get; }
        ISubMatches SubMatches { get; }
    }

    //[ComImport]
    //[Guid("3F4DACA0-160D-11D2-A8E9-00104B365C9F")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMatchCollection
    {
        int Count { get; }
        IMatch this[int index] { get; }
        System.Collections.IEnumerator GetEnumerator();
    }

    //[ComImport]
    //[Guid("3F4DACA2-160D-11D2-A8E9-00104B365C9F")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISubMatches
    {
        int Count { get; }
        object this[int index] { get; }
        System.Collections.IEnumerator GetEnumerator();
    }
}
