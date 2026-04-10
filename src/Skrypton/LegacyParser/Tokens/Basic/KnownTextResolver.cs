using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Transactions;
using Skrypton.RuntimeSupport;

namespace Skrypton.LegacyParser.Tokens.Basic;

internal static class KnownTextResolver
{

    // =======================================================================================
    // [PRIVATE] CONTENT DETERMINATION - eg. isOperator
    // =======================================================================================
    /// <summary>
    /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
    /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an codeExpression
    /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
    /// </summary>
    /// public static IEnumerable<string> ArithmeticAndStringOperatorTokenValues = new List<string>
    /// {
    /// 	"^", "/", "\\", "*", "\"", "MOD", "+", "-", "&" // Note: "\" is integer division (see the link above)
    /// }.AsReadOnly();
    // =======================================================================================
    // [PRIVATE] CONTENT DETERMINATION - eg. isOperator
    // =======================================================================================
    /// <summary>
    /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
    /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an codeExpression
    /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
    /// </summary>
    public static KnownTextContent[] k_ArithmeticAndStringOperatorTokenValues = CollectArithmeticAndStringOperatorTokenValues(); private static KnownTextContent[] CollectArithmeticAndStringOperatorTokenValues() => new KnownTextContent[]
    {
        new KnownTextContent("^", false, false, OperatorKind.Exponentiation),
        new KnownTextContent("/", false, false, OperatorKind.Division),
        new KnownTextContent("\\", false, false, OperatorKind.IntegerDivision),
        new KnownTextContent("*", false, false, OperatorKind.Multiplication ),
        new KnownTextContent("\"", false, false, OperatorKind.VerySpecial),
        new KnownTextContent("MOD", false, false, OperatorKind.Modulus),
        new KnownTextContent("+", false, false, OperatorKind.Plus),
        new KnownTextContent("-", false, false, OperatorKind.Minus),
        new KnownTextContent("&", false, false, OperatorKind.StringConcatenation), // Note: "\" is integer division (see the link above)
    };

    /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
    /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an codeExpression
    /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
    /// public static IEnumerable<string> LogicalOperatorTokenValues = new List<string>
    /// {
    /// 	"NOT", "AND", "OR", "XOR"
    /// }.AsReadOnly();
    /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
    /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an codeExpression
    /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
    public static KnownTextContent[] k_LogicalOperatorTokenValues = CollectLogicalOperatorTokenValues(); private static KnownTextContent[] CollectLogicalOperatorTokenValues() => new KnownTextContent[]
    {
        new KnownTextContent("NOT", false, false, OperatorKind.LogicalNot),
        new KnownTextContent("AND", false, false, OperatorKind.LogicalAnd),
        new KnownTextContent("OR", false, false, OperatorKind.LogicalOr),
        new KnownTextContent("XOR", false, false, OperatorKind.LogicalXor)
    };

    /// <summary>
    /// Does the content appear to represent a VBScript operator (eg. an arithermetic operator such as "*", a logical operator such as "AND" or
    /// a comparison operator such as ">")? An exception will be raised for null, blank or whitespace-containing input.
    /// </summary>
    /// protected static bool isOperator(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		ArithmeticAndStringOperatorTokenValues.Concat(LogicalOperatorTokenValues).Concat(ComparisonTokenValues)
    /// 	);
    /// }
    private static KnownTextContent[] k_OperatorNames = CollectArithmeticAndStringOperatorTokenValues()
        .Concat(CollectLogicalOperatorTokenValues())
        .Concat(CollectComparisonTokenValues()).ToArray();
    internal static OperatorKind? isOperatorUpper(StringUpper atomContent)
    {
        //if (k_OperatorNames == null)
        //{
        //    k_OperatorNames = k_ArithmeticAndStringOperatorTokenValues
        //        .Concat(k_LogicalOperatorTokenValues)
        //        .Concat(k_ComparisonTokenValues).ToArray();
        //}
        KnownTextContent? ktc = isTypeUpper(atomContent, k_OperatorNames);
        return ktc == null ? default(OperatorKind?) : (OperatorKind)ktc.ThePayload;
    }

    /// <summary>
    /// Does the content appear to represent a VBScript operator (eg. AND)? An exception will be raised
    /// for null, blank or whitespace-containing input.
    /// </summary>
    /// protected static bool isLogicalOperator(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		LogicalOperatorTokenValues
    /// 	);
    /// }
    internal static OperatorKind? isLogicalOperatorUpper(StringUpper atomContent)
    {
        KnownTextContent? ktc = isTypeUpper(atomContent, k_LogicalOperatorTokenValues);
        return ktc == null ? default(OperatorKind?) : (OperatorKind)ktc.ThePayload;
    }

    /// <summary>
    /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
    /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an codeExpression
    /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
    /// </summary>
    /// public static IEnumerable<string> ComparisonTokenValues = new List<string>
    /// {
    /// 	"=", "<>", "<", ">", "<=", ">=", "IS",
    /// 	"EQV", "IMP"
    /// }.AsReadOnly();
    public static readonly KnownTextContent[] k_ComparisonTokenValues = CollectComparisonTokenValues(); private static KnownTextContent[] CollectComparisonTokenValues() => new KnownTextContent[]
    {
        new KnownTextContent("=", false, false, OperatorKind.Equal),
        new KnownTextContent("<>", false, false, OperatorKind.NotEqual),
        new KnownTextContent("<", false, false, OperatorKind.LessThan),
        new KnownTextContent(">", false, false, OperatorKind.GreaterThan),
        new KnownTextContent("<=", false, false, OperatorKind.LessThanOrEqual),
        new KnownTextContent(">=", false, false, OperatorKind.GreaterThanOrEqual),


        ///
        /// @lubo: The Is operator is used to determine if two variables refer to the same object. The output is either True or False.
        ///
        new KnownTextContent("IS", false, false, OperatorKind.IsSameObject),

        ///
        /// @lubo: The Eqv operator is used to perform a logical comparison on two exressions (i.e., are the two expressions identical), where the expressions are Null, or are of Boolean subtype and have a value of True or False.
        /// The 'Eqv' operator can also be used a "bitwise operator" to make a bit-by-bit comparison of two integers. If both bits in the comparison are the same (both are 0's or 1's), then a 1 is returned. Otherwise, a 0 is returned.
        ///
        new KnownTextContent("EQV", false, false, OperatorKind.LogicalEquivalence),
        ///
        /// lubo: The 'Imp' operator is used to perform a logical implication on two expressions, where the expressions are Null, or are of Boolean subtype and have a value of True or False.
        /// The Imp operator can also be used a "bitwise operator" to make a bit-by-bit comparison of two integers. If both bits in the comparison are the same (both are 0's or 1's), then a 1 is returned. If the first bit is a 0 and the second bit is a 1, then a 1 is returned. If the first bit is a 1 and the second bit is a 0, then a 0 is returned.
        new KnownTextContent("IMP", false, false, OperatorKind.LogicalImplication)
    };

    /// <summary>
    /// Does the content appear to represent a VBScript comparison? An exception will be raised
    /// for null, blank or whitespace-containing input.
    /// </summary>
    /// protected static bool isComparison(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		ComparisonTokenValues
    /// 	);
    /// }
    internal static OperatorKind? isComparisonUpper(StringUpper atomContent)
    {
        KnownTextContent? ktc = isTypeUpper(atomContent, k_ComparisonTokenValues);
        return ktc == null ? default(OperatorKind?) : (OperatorKind)ktc.ThePayload;
    }


    /// protected static bool isMemberAccessor(string atomContent)
    /// 		{
    /// 			return isType(
    /// 				atomContent,
    /// 				new string[] { "." }
    /// 			);
    /// 		}
    private static readonly KnownTextContent[] dot = new[] { new KnownTextContent(".", false, false, null) };
    internal static bool isMemberAccessorUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, dot) != null;
    }


    /// private static bool isArgumentSeparator(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		new string[] { "," }
    /// 	);
    /// }
    private static readonly KnownTextContent[] k_Separator = new[] { new KnownTextContent(",", false, false, null) };
    internal static bool isArgumentSeparatorUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_Separator) != null;
    }

    private static readonly KnownTextContent[] k_OpenBrace = new[] { new KnownTextContent("(", false, false, null) };

    internal static bool isOpenBraceUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_OpenBrace) != null;
    }

    private static readonly KnownTextContent[] k_CloseBrace = new[] { new KnownTextContent(")", false, false, null) };
    internal static bool isCloseBraceUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_CloseBrace) != null;
    }


    private static readonly KnownTextContent[] k_me = new[] { new KnownTextContent("me", false, false, null) };
    internal static bool isTargetCurrentClassTokenUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_me) != null;
    }
    /// <summary>
    /// Does the content appear to represent a VBScript keyword that will have to be handled by an
    /// AbstractBlockHandler - eg. a "FOR" in a loop, or the "OPTION" from "OPTION EXPLICIT" or the
    /// "RANDOMIZE" command? An exception will be raised for null, blank or whitespace-containing
    /// input.
    /// </summary>
    /// protected static bool isMustHandleKeyWord(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		new string[]
    /// 		{
    /// 			"OPTION",
    /// 			"DIM", "REDIM", "PRESERVE",
    /// 			"PUBLIC", "PRIVATE",
    /// 			"IF", "THEN", "ELSE", "ELSEIF", "END",
    /// 			"WITH",
    /// 			"SUB", "FUNCTION", "CLASS",
    /// 			"EXIT",
    /// 			"SELECT", "CASE",
    /// 			"FOR", "EACH", "NEXT", "TO",
    /// 			"DO", "WHILE", "UNTIL", "LOOP", "WEND",
    /// 			"RANDOMIZE",
    /// 			"REM",
    /// 			"GET",
    ///
    /// 			// This is a keyword, not a function, since built-in functions all take by-val arguments
    /// 			// while this would be an exception if it was identified as a function (relying upon all
    /// 			// built-in functions only taking by-val arguments allows for some shortcuts in the
    /// 			// translation process)
    /// 			"ERASE"
    /// 		}
    /// 	);
    /// }
    ///
    private static KnownTextContent[] StringContentCollectionCreate(string[] src)
    {
        return src.Select(x => new KnownTextContent(x, false, false, null)).ToArray();
    }
    /// <summary>
    /// Does the content appear to represent a VBScript keyword that will have to be handled by an
    /// AbstractBlockHandler - eg. a "FOR" in a loop, or the "OPTION" from "OPTION EXPLICIT" or the
    /// "RANDOMIZE" command? An exception will be raised for null, blank or whitespace-containing
    /// input.
    /// </summary>
    private static readonly KnownTextContent[] k_MustHandleKeyWord = StringContentCollectionCreate(
        new string[]
        {
            "OPTION",
            "DIM", "REDIM", "PRESERVE",
            "PUBLIC", "PRIVATE",
            "IF", "THEN", "ELSE", "ELSEIF", "END",
            "WITH",
            "SUB", "FUNCTION", "CLASS",
            "EXIT",
            "SELECT", "CASE",
            "FOR", "EACH", "NEXT", "TO",
            "DO", "WHILE", "UNTIL", "LOOP", "WEND",
            "RANDOMIZE",
            "REM",
            "GET",

            // This is a keyword, not a function, since built-in functions all take by-val arguments
            // while this would be an exception if it was identified as a function (relying upon all
            // built-in functions only taking by-val arguments allows for some shortcuts in the
            // translation process)
            "ERASE"
        }

    ).ToArray();

    internal static bool isMustHandleKeyWordUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_MustHandleKeyWord) != null;
    }
    /// <summary>
    /// There are some keywords that can be used as variable names, deeper parsing work than looking
    /// solely at its content is required to determine if a token is a NameToken or KeywordToken
    /// </summary>
    /// protected static bool isContextDependentKeyword(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		new string[]
    /// 		{
    /// 			"EXPLICIT", "PROPERTY", "DEFAULT", "STEP", "ERROR"
    /// 		}
    /// 	);
    /// }
    private static readonly KnownTextContent[] k_ContextDependentKeyword = StringContentCollectionCreate(
        new string[]
        {
            "EXPLICIT", "PROPERTY", "DEFAULT", "STEP", "ERROR"
        }

    ).ToArray();
    internal static bool isContextDependentKeywordUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_ContextDependentKeyword) != null;
    }


    /// <summary>
    /// Does the content appear to represent a VBScript keyword that may form part of a general
    /// statement and not have to be handled by a specific AbstractBlockHandler - eg. a "NEW"
    /// declaration for instantiating a class instance. An exception will be raised for null,
    /// blank or whitespace-containing input.
    /// </summary>
    /// protected static bool isMiscKeyWord(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		new string[]
    /// 		{
    /// 			"CALL",
    /// 			"LET", "SET",
    /// 			"NEW",
    /// 			"ON", "RESUME"
    /// 		}
    /// 	);
    /// }
    private static readonly KnownTextContent[] k_MiscKeyWord = StringContentCollectionCreate(
        new string[]
        {
            "CALL",
            "LET", "SET",
            "NEW",
            "ON", "RESUME"
        }
    ).ToArray();

    internal static bool isMiscKeyWordUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_MiscKeyWord) != null;
    }


    /// <summary>
    /// Does the content appear to represent a VBScript codeExpression - eg. "NOTHING". An exception will be raised for null, blank or whitespace-containing input.
    /// </summary>
    /// protected static bool isVBScriptValue(string atomContent)
    /// {
    /// 	return isType(
    /// 		atomContent,
    /// 		new string[]
    /// 		{
    /// 			"TRUE", "FALSE",
    /// 			"EMPTY", "NOTHING", "NULL",
    /// 			"ERR",
    ///
    /// 			// These are the constants from http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbscon3.htm that appear to work in VBScript
    ///
    /// 			// VarType Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs57.htm)
    /// 			"vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
    /// 			"vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbArray",
    ///
    /// 			// MsgBox Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs49.htm) - don't know why these are defined, but they are!
    /// 			"vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel", "vbYesNo", "vbRetryCancel", "vbCritical", "vbQuestion", "vbExclamation",
    /// 			"vbInformation", "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", "vbDefaultButton4", "vbApplicationModal", "vbSystemModal",
    /// 			"vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",
    ///
    /// 			// String Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs53.htm)
    /// 			"vbCr", "vbCrLf", "vbFormFeed", "vbLf", "vbNewLine", "vbNullChar", "vbNullString", "vbTab", "vbVerticalTab",
    ///
    /// 			// Miscellaneous Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs47.htm)
    /// 			"vbObjectError",
    ///
    /// 			// Comparison Constants  (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs35.htm)
    /// 			"vbBinaryCompare", "vbTextCompare",
    ///
    /// 			// Date and Time Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs39.htm)
    /// 			"vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday", "vbFriday", "vbSaturday", "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
    /// 			"vbUseSystem", "vbUseSystemDayOfWeek",
    ///
    /// 			// Colour Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs33.htm)
    /// 			"vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta", "vbCyan", "vbWhite",
    ///
    /// 			// Date Format Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs37.htm)
    /// 			"vbGeneralDate", "vbLongDate", "vbShortDate", "vbLongTime", "vbShortTime"
    /// 		}
    /// 	);
    /// }
    /// <summary>
    /// Does the content appear to represent a VBScript codeExpression - eg. "NOTHING". An exception will be raised for null, blank or whitespace-containing input.
    /// </summary>
    private static readonly KnownTextContent[] k_VBScriptValue = StringContentCollectionCreate(
        new string[]
        {
            "TRUE", "FALSE",
            "EMPTY", "NOTHING", "NULL",
            "ERR",

            // These are the constants from http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbscon3.htm that appear to work in VBScript

            // VarType Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs57.htm)
            "vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
            "vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbArray",

            // MsgBox Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs49.htm) - don't know why these are defined, but they are!
            "vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel", "vbYesNo", "vbRetryCancel", "vbCritical", "vbQuestion", "vbExclamation",
            "vbInformation", "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", "vbDefaultButton4", "vbApplicationModal", "vbSystemModal",
            "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",

            // String Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs53.htm)
            "vbCr", "vbCrLf", "vbFormFeed", "vbLf", "vbNewLine", "vbNullChar", "vbNullString", "vbTab", "vbVerticalTab",

            // Miscellaneous Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs47.htm)
            "vbObjectError",

            // Comparison Constants  (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs35.htm)
            "vbBinaryCompare", "vbTextCompare",

            // Date and Time Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs39.htm)
            "vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday", "vbFriday", "vbSaturday", "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
            "vbUseSystem", "vbUseSystemDayOfWeek",

            // Colour Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs33.htm)
            "vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta", "vbCyan", "vbWhite",

            // Date Format Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs37.htm)
            "vbGeneralDate", "vbLongDate", "vbShortDate", "vbLongTime", "vbShortTime"
        }
    ).ToArray();

    internal static bool isVBScriptValueUpper(StringUpper atomContent)
    {
        return isTypeUpper(atomContent, k_VBScriptValue) != null;
    }


    /// <summary>
    /// Does the content appear to represent a VBScript function - eg. the "ISNULL" method.
    /// An exception will be raised for null, blank or whitespace-containing input.
    /// </summary>
    /*private static readonly KnownTextContent[] xk_VBScriptFunction = StringContentCollectionCreate(
        new string[]
        {
            // Note: Some of these functions sound like they would be returned by isVBScriptFunctionThatAlwaysReturnsNumericContent but they
            // return null in some cases and so are not applicable - eg. "INT" will return Null if Null is passed in
            "ISEMPTY", "ISNULL", "ISOBJECT", "ISNUMERIC", "ISDATE", "ISARRAY",
            "VARTYPE", "TYPENAME",
            "CREATEOBJECT", "GETOBJECT",
            "CBOOL", "CSTR",
            "DATEVALUE", "TIMEVALUE",
            "DAY", "MONTH", "MONTHNAME", "YEAR", "WEEKDAY", "WEEKDAYNAME", "HOUR", "MINUTE", "SECOND", "DATEDIFF", "DATEPART",
            "ABS",
            "HEX", "OCT", "FIX",
            "INT",
            "CHR", "CHRB", "CHRW",
            "INSTR", "INSTRREV",
            "LEN", "LENB",
            "LCASE", "UCASE",
            "LEFT", "LEFTB", "RIGHT", "RIGHTB", "SPACE",
            "REPLACE",
            "STRCOMP", "STRING",
            "LTRIM", "RTRIM", "TRIM",
            "SPLIT", "ARRAY", "JOIN",
            "EVAL", "EXECUTE", "EXECUTEGLOBAL",
            "FORMATCURRENCY", "FORMATDATETIME", "FORMATNUMBER", "FORMATPERCENT",
            "FILTER", "GETLOCALE", "GETREF", "INPUTBOX", "LOADPICTURE", "MID", "MSGBOX", "RGB", "SETLOCALE", "SGN", "STRREVERSE",
            "SCRIPTENGINE",
            "ESCAPE", "UNESCAPE"
        }
    ).ToArray();*/
    //private static readonly KnownTextContent[] k_VBScriptFunction = InitializeBuiltInFunctionRegistration().Contents;
    private static readonly BuiltInFunctionRegistration k_VBScriptFunction = InitializeBuiltInFunctionRegistration();
    private static BuiltInFunctionRegistration InitializeBuiltInFunctionRegistration()
    {
        // Note: Some of these functions sound like they would be returned by isVBScriptFunctionThatAlwaysReturnsNumericContent but they
        // return null in some cases and so are not applicable - eg. "INT" will return Null if Null is passed in
        return new BuiltInFunctionRegistration()
        .AddS(BuiltInFunctionId.BuiltInFunctionISEMPTY, "ISEMPTY")
        .AddS(BuiltInFunctionId.BuiltInFunctionISNULL, "ISNULL")
        .AddS(BuiltInFunctionId.BuiltInFunctionISOBJECT, "ISOBJECT")
        .AddS(BuiltInFunctionId.BuiltInFunctionISNUMERIC, "ISNUMERIC")
        .AddS(BuiltInFunctionId.BuiltInFunctionISDATE, "ISDATE")
        .AddS(BuiltInFunctionId.BuiltInFunctionISARRAY, "ISARRAY")
        .AddS(BuiltInFunctionId.BuiltInFunctionVARTYPE, "VARTYPE")
        .AddS(BuiltInFunctionId.BuiltInFunctionTYPENAME, "TYPENAME")
        .AddS(BuiltInFunctionId.BuiltInFunctionCREATEOBJECT, "CREATEOBJECT")
        .AddS(BuiltInFunctionId.BuiltInFunctionGETOBJECT, nameof(IProvideVBScriptCompatFunctionalityToIndividualRequests.GETOBJECT))
        .AddS(BuiltInFunctionId.BuiltInFunctionCBOOL, "CBOOL")
        .AddS(BuiltInFunctionId.BuiltInFunctionCSTR, "CSTR")
        .AddS(BuiltInFunctionId.BuiltInFunctionDATEVALUE, "DATEVALUE")
        .AddS(BuiltInFunctionId.BuiltInFunctionTIMEVALUE, "TIMEVALUE")
        .AddS(BuiltInFunctionId.BuiltInFunctionDAY, "DAY")
        .AddS(BuiltInFunctionId.BuiltInFunctionMONTH, "MONTH")
        .AddS(BuiltInFunctionId.BuiltInFunctionMONTHNAME, "MONTHNAME")
        .AddS(BuiltInFunctionId.BuiltInFunctionYEAR, "YEAR")
        .AddS(BuiltInFunctionId.BuiltInFunctionWEEKDAY, "WEEKDAY")
        .AddS(BuiltInFunctionId.BuiltInFunctionWEEKDAYNAME, "WEEKDAYNAME")
        .AddS(BuiltInFunctionId.BuiltInFunctionHOUR, "HOUR")
        .AddS(BuiltInFunctionId.BuiltInFunctionMINUTE, "MINUTE")
        .AddS(BuiltInFunctionId.BuiltInFunctionSECOND, "SECOND")
        .AddS(BuiltInFunctionId.BuiltInFunctionDATEDIFF, "DATEDIFF")
        .AddS(BuiltInFunctionId.BuiltInFunctionDATEPART, "DATEPART")
        .AddS(BuiltInFunctionId.BuiltInFunctionABS, "ABS")
        //.AddS(BuiltInFunctionId.BuiltInFunctionHEX, "HEX") numeric
        //.AddS(BuiltInFunctionId.BuiltInFunctionOCT, "OCT") numeric
        //.AddS(BuiltInFunctionId.BuiltInFunctionFIX, "FIX") numeric
        //.AddS(BuiltInFunctionId.BuiltInFunctionINT, "INT") numeric
        .AddS(BuiltInFunctionId.BuiltInFunctionCHR, "CHR")
        .AddS(BuiltInFunctionId.BuiltInFunctionCHRB, "CHRB")
        .AddS(BuiltInFunctionId.BuiltInFunctionCHRW, "CHRW")
        .AddS(BuiltInFunctionId.BuiltInFunctionINSTR, "INSTR")
        .AddS(BuiltInFunctionId.BuiltInFunctionINSTRREV, "INSTRREV")
        .AddS(BuiltInFunctionId.BuiltInFunctionLEN, "LEN")
        .AddS(BuiltInFunctionId.BuiltInFunctionLENB, "LENB")
        .AddS(BuiltInFunctionId.BuiltInFunctionLCASE, "LCASE")
        .AddS(BuiltInFunctionId.BuiltInFunctionUCASE, "UCASE")
        .AddS(BuiltInFunctionId.BuiltInFunctionLEFT, "LEFT")
        .AddS(BuiltInFunctionId.BuiltInFunctionLEFTB, "LEFTB")
        .AddS(BuiltInFunctionId.BuiltInFunctionRIGHT, "RIGHT")
        .AddS(BuiltInFunctionId.BuiltInFunctionRIGHTB, "RIGHTB")
        .AddS(BuiltInFunctionId.BuiltInFunctionSPACE, "SPACE")
        .AddS(BuiltInFunctionId.BuiltInFunctionREPLACE, "REPLACE")
        .AddS(BuiltInFunctionId.BuiltInFunctionSTRCOMP, "STRCOMP")
        .AddS(BuiltInFunctionId.BuiltInFunctionSTRING, "STRING")
        .AddS(BuiltInFunctionId.BuiltInFunctionLTRIM, "LTRIM")
        .AddS(BuiltInFunctionId.BuiltInFunctionRTRIM, "RTRIM")
        .AddS(BuiltInFunctionId.BuiltInFunctionTRIM, "TRIM")
        .AddS(BuiltInFunctionId.BuiltInFunctionSPLIT, "SPLIT")
        .AddS(BuiltInFunctionId.BuiltInFunctionARRAY, "ARRAY")
        .AddS(BuiltInFunctionId.BuiltInFunctionJOIN, "JOIN")
        .AddS(BuiltInFunctionId.BuiltInFunctionEVAL, "EVAL")
        .AddS(BuiltInFunctionId.BuiltInFunctionEXECUTE, "EXECUTE")
        .AddS(BuiltInFunctionId.BuiltInFunctionEXECUTEGLOBAL, "EXECUTEGLOBAL")
        .AddS(BuiltInFunctionId.BuiltInFunctionFORMATCURRENCY, "FORMATCURRENCY")
        .AddS(BuiltInFunctionId.BuiltInFunctionFORMATDATETIME, "FORMATDATETIME")
        .AddS(BuiltInFunctionId.BuiltInFunctionFORMATNUMBER, "FORMATNUMBER")
        .AddS(BuiltInFunctionId.BuiltInFunctionFORMATPERCENT, "FORMATPERCENT")
        .AddS(BuiltInFunctionId.BuiltInFunctionFILTER, "FILTER")
        .AddS(BuiltInFunctionId.BuiltInFunctionGETLOCALE, "GETLOCALE")
        .AddS(BuiltInFunctionId.BuiltInFunctionGETREF, "GETREF")
        .AddS(BuiltInFunctionId.BuiltInFunctionINPUTBOX, "INPUTBOX")
        .AddS(BuiltInFunctionId.BuiltInFunctionLOADPICTURE, "LOADPICTURE")
        .AddS(BuiltInFunctionId.BuiltInFunctionMID, "MID")
        .AddS(BuiltInFunctionId.BuiltInFunctionMSGBOX, "MSGBOX")
        .AddS(BuiltInFunctionId.BuiltInFunctionRGB, "RGB")
        .AddS(BuiltInFunctionId.BuiltInFunctionSETLOCALE, "SETLOCALE")
        .AddS(BuiltInFunctionId.BuiltInFunctionSGN, "SGN")
        .AddS(BuiltInFunctionId.BuiltInFunctionSTRREVERSE, "STRREVERSE")
        .AddS(BuiltInFunctionId.BuiltInFunctionSCRIPTENGINE, "SCRIPTENGINE")
        .AddS(BuiltInFunctionId.BuiltInFunctionESCAPE, "ESCAPE")
        .AddS(BuiltInFunctionId.BuiltInFunctionUNESCAPE, "UNESCAPE")
        //
        .AddN(BuiltInFunctionId.BuiltInFunctionLBOUND, "LBOUND")
        .AddN(BuiltInFunctionId.BuiltInFunctionUBOUND, "UBOUND")
        .AddN(BuiltInFunctionId.BuiltInFunctionCBYTE, "CBYTE")
        .AddN(BuiltInFunctionId.BuiltInFunctionCCUR, "CCUR")
        .AddN(BuiltInFunctionId.BuiltInFunctionCINT, "CINT")
        .AddN(BuiltInFunctionId.BuiltInFunctionCLNG, "CLNG")
        .AddN(BuiltInFunctionId.BuiltInFunctionCSNG, "CSNG")
        .AddN(BuiltInFunctionId.BuiltInFunctionCDBL, "CDBL")
        .AddN(BuiltInFunctionId.BuiltInFunctionCDATE, "CDATE")
        .AddN(BuiltInFunctionId.BuiltInFunctionDATEADD, "DATEADD")
        .AddN(BuiltInFunctionId.BuiltInFunctionDATESERIAL, "DATESERIAL")
        .AddN(BuiltInFunctionId.BuiltInFunctionTIMESERIAL, "TIMESERIAL")
        .AddN(BuiltInFunctionId.BuiltInFunctionNOW, "NOW")
        .AddN(BuiltInFunctionId.BuiltInFunctionDATE, "DATE")
        .AddN(BuiltInFunctionId.BuiltInFunctionTIME, "TIME")
        .AddN(BuiltInFunctionId.BuiltInFunctionATN, "ATN")
        .AddN(BuiltInFunctionId.BuiltInFunctionCOS, "COS")
        .AddN(BuiltInFunctionId.BuiltInFunctionSIN, "SIN")
        .AddN(BuiltInFunctionId.BuiltInFunctionTAN, "TAN")
        .AddN(BuiltInFunctionId.BuiltInFunctionEXP, "EXP")
        .AddN(BuiltInFunctionId.BuiltInFunctionLOG, "LOG")
        .AddN(BuiltInFunctionId.BuiltInFunctionSQR, "SQR")
        .AddN(BuiltInFunctionId.BuiltInFunctionRND, "RND")
        .AddN(BuiltInFunctionId.BuiltInFunctionROUND, "ROUND")
        .AddN(BuiltInFunctionId.BuiltInFunctionHEX, "HEX")
        .AddN(BuiltInFunctionId.BuiltInFunctionOCT, "OCT")
        .AddN(BuiltInFunctionId.BuiltInFunctionFIX, "FIX")
        .AddN(BuiltInFunctionId.BuiltInFunctionINT, "INT")
        .AddN(BuiltInFunctionId.BuiltInFunctionSNG, "SNG")
        .AddN(BuiltInFunctionId.BuiltInFunctionASC, "ASC")
        .AddN(BuiltInFunctionId.BuiltInFunctionASCB, "ASCB")
        .AddN(BuiltInFunctionId.BuiltInFunctionASCW, "ASCW")
        .AddN(BuiltInFunctionId.BuiltInFunctionSCRIPTENGINEBUILDVERSION, "SCRIPTENGINEBUILDVERSION")
        .AddN(BuiltInFunctionId.BuiltInFunctionSCRIPTENGINEMAJORVERSION, "SCRIPTENGINEMAJORVERSION")
        .AddN(BuiltInFunctionId.BuiltInFunctionSCRIPTENGINEMINORVERSION, "SCRIPTENGINEMINORVERSION")
        .AddN(BuiltInFunctionId.BuiltInFunctionTIMER, "TIMER")
        .Build();
    }

    internal static bool isVBScriptFunctionUpper(StringUpper atomContent, out BuiltInFunctionInfo? functionInfo)
    {
        if (isTypeUpperF(atomContent, k_VBScriptFunction.FunctionInfos, out functionInfo))
        {
            return true;
        }
        return false;
    }

    private sealed class BuiltInFunctionRegistration
    {
        private readonly Dictionary<string, BuiltInFunctionInfo> _registry = new Dictionary<string, BuiltInFunctionInfo>();
        internal IReadOnlyDictionary<string, BuiltInFunctionInfo> FunctionInfos { get; private set; } = new Dictionary<string, BuiltInFunctionInfo>();
        internal KnownTextContent[] FunctionContents { get; private set; } = [];

        public BuiltInFunctionRegistration AddN(BuiltInFunctionId functionId, string name) => AddS(functionId, name, alwaysReturnsNumeric: true);
        public BuiltInFunctionRegistration AddS(BuiltInFunctionId functionId, string name, bool alwaysReturnsNumeric = false)
        {
            if ((!Enum.TryParse<BuiltInFunctionId>("BuiltInFunction" + name, out BuiltInFunctionId enumItem) || (enumItem != functionId)))
            {
                throw new InvalidOperationException($"The functionId {functionId} does not match the expected enum item name of BuiltInFunction{name}");
            }
            _registry.Add(name, new BuiltInFunctionInfo(functionId, alwaysReturnsNumeric, new KnownTextContent(name, false, false, null)));
            return this;
        }

        public BuiltInFunctionRegistration Build()
        {
            FunctionInfos = _registry.ToDictionary(x => x.Key, x => x.Value);
            FunctionContents = FunctionInfos.Values.Select(x => x.Content).ToArray();

            return this;
        }
    }
    /// <summary>
    /// Does the content appear to represent a VBScript function that is guaranteed to return a numeric value - eg. the "CDBL" method. This does
    /// not include functions that will ever return Null (such as ABS, which returns Null if Null is provided as the argument). An exception will
    /// be raised for null, blank or whitespace-containing input.
    /*
    /// </summary>
    private static readonly KnownTextContent[] x_VBScriptFunctionThatAlwaysReturnsNumericContent = StringContentCollectionCreate(
        new string[]
        {
            "LBOUND", "UBOUND",
            "CBYTE", "CCUR", "CINT", "CLNG", "CSNG", "CDBL", "CDATE",
            "DATEADD", "DATESERIAL", "TIMESERIAL",
            "NOW", "DATE", "TIME",
            "ATN", "COS", "SIN", "TAN", "EXP", "LOG", "SQR", "RND", "ROUND",
            "HEX", "OCT", "FIX", "INT", "SNG",
            "ASC", "ASCB", "ASCW",
            "SCRIPTENGINEBUILDVERSION", "SCRIPTENGINEMAJORVERSION", "SCRIPTENGINEMINORVERSION",
            "TIMER"
        }
    ).ToArray();
*/

    internal static bool containsWhiteSpace(string content)
    {
        if (content == null)
            throw new ArgumentNullException(nameof(content));

        for (int ix = 0; ix < content.Length; ix++)
        {
            if (char.IsWhiteSpace(content[ix]))
                return true;
        }

        return false;
        //return content.Any(c => WhiteSpaceChars.IndexOf(c) != -1);
    }

    private static KnownTextContent? isTypeUpper(StringUpper atomContent, KnownTextContent[] keyWords)
    {
        if (atomContent == null)
        {
            throw new ArgumentNullException(nameof(atomContent));
        }

        if (atomContent.Length == 0)
        {
            throw new ArgumentException("Blank content specified - invalid");
        }

        if (keyWords == null)
        {
            throw new ArgumentNullException(nameof(keyWords));
        }

        for (int ix = 0; ix < keyWords.Length; ix++)
        {
            KnownTextContent keyWord = keyWords[ix];

            if (keyWord.EqualsCaseUpper(atomContent))
            {
                return keyWord;
            }
        }

        if (atomContent.containsWhiteSpace())
        {
            throw new ArgumentException("Whitespace encountered in atomContent - invalid");
        }
        return null;
    }
    private static bool isTypeUpperF(StringUpper atomContent, IReadOnlyDictionary<string, BuiltInFunctionInfo> keyWords, out BuiltInFunctionInfo? result)
    {
        if (atomContent == null)
        {
            throw new ArgumentNullException(nameof(atomContent));
        }

        if (atomContent.Length == 0)
        {
            throw new ArgumentException("Blank content specified - invalid");
        }

        if (keyWords == null)
        {
            throw new ArgumentNullException(nameof(keyWords));
        }

        foreach (var kw in keyWords)
        {
            if (kw.Value.Content.EqualsCaseUpper(atomContent))
            {
                result = kw.Value;
                return true;
            }
        }

        if (atomContent.containsWhiteSpace())
        {
            throw new ArgumentException("Whitespace encountered in atomContent - invalid");
        }
        result = default(BuiltInFunctionInfo);
        return false;
    }
}


public sealed class KnownTextContent
{
    //private bool? isContainsWhiteSpace;
    //private bool? xIsNullOrBlank;
    /// public StringContent(string content)
    /// {
    /// TheContent = content;
    /// }
    private readonly object? thePayload;
    private readonly int length;
    public KnownTextContent(string content, bool isWhiteSpace, bool isNull, object? payload)
    {
        if (isWhiteSpace || isNull)
        {
            throw new NotSupportedException();
        }

        TheContent = content;
        TheContentUpper = content.ToUpperX();
        theContentUpperHashCode = TheContentUpper.GetHashCode();
        thePayload = payload;
        theContentUpper = content.ToUpperInvariant();
        //isContainsWhiteSpace = isWhiteSpace;
        //xIsNullOrBlank = isNull;
        length = content.Length;
    }
    internal object ThePayload
    {
        get
        {
            if (thePayload == null)
            {
                throw new NotSupportedException("No payload for " + TheContent);
            }

            return thePayload;
        }
    }

    private readonly string theContentUpper;
    private readonly int theContentUpperHashCode;
    public string TheContent { get; }
    public StringUpper TheContentUpper { get; }

    //internal bool ContainsWhiteSpace()
    //{
    //    ///if (!isContainsWhiteSpace.HasValue)
    //    ///{
    //    ///    isContainsWhiteSpace = AtomToken.containsWhiteSpace(TheContent);
    //    ///}
    //    ///return isContainsWhiteSpace.Value;
    //    return false;
    //}

    /// internal bool EqualsIgnoreCase(string atomContent)
    /// {
    /// return string.Equals(TheContent, atomContent, StringComparison.OrdinalIgnoreCase);
    /// }
    internal bool EqualsCaseUpper(StringUpper atomContent)
    {
        if (this.length != atomContent.Length)
            return false;
        var hc = atomContent.UpperText.GetHashCode(StringComparison.Ordinal);
        if (this.theContentUpperHashCode == hc)
        {
            if (!string.Equals(theContentUpper, atomContent.UpperText, StringComparison.Ordinal))
                throw new NotImplementedException(atomContent.Original + " # " + this.TheContent);
            return true;
        }
        /// if (string.Equals(theContentUpper, atomContent.UpperText, StringComparison.Ordinal))
        /// throw new NotImplementedException(atomContent.Original + " # " + this.TheContent);
        /// return false;
        return string.Equals(theContentUpper, atomContent.UpperText, StringComparison.Ordinal);
    }

    //internal bool IsNullOrBlank()
    //{
    //    /// if (!xIsNullOrBlank.HasValue)
    //    /// {
    //    /// xIsNullOrBlank = ((TheContent ?? "").Trim() == "");
    //    /// }
    //    /// return xIsNullOrBlank.Value;
    //    return false;
    //}
}

[DebuggerDisplay("{UpperText}")]
public sealed class StringUpper
{
    internal readonly string Original;
    internal readonly string UpperText;
    internal readonly int Length;
    internal bool? hasWhiteSpace;
    public StringUpper(string original)
    {
        this.Original = original ?? throw new ArgumentNullException(nameof(original));
        this.UpperText = Original.ToUpperInvariant();
        this.Length = original.Length;
    }

    internal bool IsNullOrWhiteSpace()
    {
        return Length == 0;
    }

    internal bool containsWhiteSpace()
    {
        if (!hasWhiteSpace.HasValue)
        {
            hasWhiteSpace = KnownTextResolver.containsWhiteSpace(UpperText);
        }
        return hasWhiteSpace.Value;
    }
}

static class StringExtensionUpper
{
    internal static StringUpper ToUpperX(this string content)
    {
        if (content == null)
            throw new ArgumentNullException(nameof(content));

        return new StringUpper(content);
    }
}

#pragma warning disable SA1602 // Enumeration items should be documented
internal enum BuiltInFunctionId
{
    Unknown,
    BuiltInFunctionISEMPTY,
    BuiltInFunctionISNULL,
    BuiltInFunctionISOBJECT,
    BuiltInFunctionISNUMERIC,
    BuiltInFunctionISDATE,
    BuiltInFunctionISARRAY,
    BuiltInFunctionVARTYPE,
    BuiltInFunctionTYPENAME,
    BuiltInFunctionCREATEOBJECT,
    BuiltInFunctionGETOBJECT,
    BuiltInFunctionCBOOL,
    BuiltInFunctionCSTR,
    BuiltInFunctionDATEVALUE,
    BuiltInFunctionTIMEVALUE,
    BuiltInFunctionDAY,
    BuiltInFunctionMONTH,
    BuiltInFunctionMONTHNAME,
    BuiltInFunctionYEAR,
    BuiltInFunctionWEEKDAY,
    BuiltInFunctionWEEKDAYNAME,
    BuiltInFunctionHOUR,
    BuiltInFunctionMINUTE,
    BuiltInFunctionSECOND,
    BuiltInFunctionDATEDIFF,
    BuiltInFunctionDATEPART,
    BuiltInFunctionABS,
    //BuiltInFunctionHEX, numeric
    //BuiltInFunctionOCT, numeric
    //BuiltInFunctionFIX, numeric
    //BuiltInFunctionINT, numeric
    BuiltInFunctionCHR,
    BuiltInFunctionCHRB,
    BuiltInFunctionCHRW,
    BuiltInFunctionINSTR,
    BuiltInFunctionINSTRREV,
    BuiltInFunctionLEN,
    BuiltInFunctionLENB,
    BuiltInFunctionLCASE,
    BuiltInFunctionUCASE,
    BuiltInFunctionLEFT,
    BuiltInFunctionLEFTB,
    BuiltInFunctionRIGHT,
    BuiltInFunctionRIGHTB,
    BuiltInFunctionSPACE,
    BuiltInFunctionREPLACE,
    BuiltInFunctionSTRCOMP,
    BuiltInFunctionSTRING,
    BuiltInFunctionLTRIM,
    BuiltInFunctionRTRIM,
    BuiltInFunctionTRIM,
    BuiltInFunctionSPLIT,
    BuiltInFunctionARRAY,
    BuiltInFunctionJOIN,
    BuiltInFunctionEVAL,
    BuiltInFunctionEXECUTE,
    BuiltInFunctionEXECUTEGLOBAL,
    BuiltInFunctionFORMATCURRENCY,
    BuiltInFunctionFORMATDATETIME,
    BuiltInFunctionFORMATNUMBER,
    BuiltInFunctionFORMATPERCENT,
    BuiltInFunctionFILTER,
    BuiltInFunctionGETLOCALE,
    BuiltInFunctionGETREF,
    BuiltInFunctionINPUTBOX,
    BuiltInFunctionLOADPICTURE,
    BuiltInFunctionMID,
    BuiltInFunctionMSGBOX,
    BuiltInFunctionRGB,
    BuiltInFunctionSETLOCALE,
    BuiltInFunctionSGN,
    BuiltInFunctionSTRREVERSE,
    BuiltInFunctionSCRIPTENGINE,
    BuiltInFunctionESCAPE,
    BuiltInFunctionUNESCAPE,

    //
    BuiltInFunctionLBOUND,
    BuiltInFunctionUBOUND,
    BuiltInFunctionCBYTE,
    BuiltInFunctionCCUR,
    BuiltInFunctionCINT,
    BuiltInFunctionCLNG,
    BuiltInFunctionCSNG,
    BuiltInFunctionCDBL,
    BuiltInFunctionCDATE,
    BuiltInFunctionDATEADD,
    BuiltInFunctionDATESERIAL,
    BuiltInFunctionTIMESERIAL,
    BuiltInFunctionNOW,
    BuiltInFunctionDATE,
    BuiltInFunctionTIME,
    BuiltInFunctionATN,
    BuiltInFunctionCOS,
    BuiltInFunctionSIN,
    BuiltInFunctionTAN,
    BuiltInFunctionEXP,
    BuiltInFunctionLOG,
    BuiltInFunctionSQR,
    BuiltInFunctionRND,
    BuiltInFunctionROUND,
    BuiltInFunctionHEX,
    BuiltInFunctionOCT,
    BuiltInFunctionFIX,
    BuiltInFunctionINT,
    BuiltInFunctionSNG,
    BuiltInFunctionASC,
    BuiltInFunctionASCB,
    BuiltInFunctionASCW,
    BuiltInFunctionSCRIPTENGINEBUILDVERSION,
    BuiltInFunctionSCRIPTENGINEMAJORVERSION,
    BuiltInFunctionSCRIPTENGINEMINORVERSION,
    BuiltInFunctionTIMER,
}
#pragma warning restore SA1602 // Enumeration items should be documented
[DebuggerDisplay("{Content.TheContent} {FunctionId}")]
internal sealed class BuiltInFunctionInfo
{
    public BuiltInFunctionId FunctionId { get; }

    /// Does the content appear to represent a VBScript function that is guaranteed to return a numeric value - eg. the "CDBL" method. This does
    /// not include functions that will ever return Null (such as ABS, which returns Null if Null is provided as the argument). An exception will
    /// be raised for null, blank or whitespace-containing input.
    public bool AlwaysReturnsNumeric { get; } // These must ONLY include those that will never return null
    public KnownTextContent Content { get; }

    public BuiltInFunctionInfo(BuiltInFunctionId functionId, bool alwaysReturnsNumeric, KnownTextContent content)
    {
        FunctionId = functionId;
        AlwaysReturnsNumeric = alwaysReturnsNumeric;
        Content = content;
    }
}