using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// The Err.Number, Err.Description and Err.Source of an exception, as Windows VBScript has them (issue #85): the runtime errors of VBScript
    /// have their standard descriptions and the source "Microsoft VBScript runtime error", Err.Raise keeps its description and source, .NET
    /// exceptions are mapped to the VBScript errors they correspond to
    /// </summary>
    public sealed class VBScriptErrorInfo
    {
        public const string RuntimeErrorSource = "Microsoft VBScript runtime error";

        /// <summary>
        /// The descriptions of the VBScript runtime errors (as Windows VBScript reports them)
        /// </summary>
        private static readonly Dictionary<int, string> Descriptions = new Dictionary<int, string>
        {
            { 3, "Return without GoSub" }, { 5, "Invalid procedure call or argument" }, { 6, "Overflow" }, { 7, "Out of memory" },
            { 9, "Subscript out of range" }, { 10, "This array is fixed or temporarily locked" }, { 11, "Division by zero" }, { 13, "Type mismatch" },
            { 14, "Out of string space" }, { 17, "Can't perform requested operation" }, { 28, "Out of stack space" }, { 35, "Sub or Function not defined" },
            { 48, "Error in loading DLL" }, { 51, "Internal error" }, { 52, "Bad file name or number" }, { 53, "File not found" },
            { 54, "Bad file mode" }, { 55, "File already open" }, { 57, "Device I/O error" }, { 58, "File already exists" },
            { 61, "Disk full" }, { 62, "Input past end of file" }, { 67, "Too many files" }, { 68, "Device unavailable" },
            { 70, "Permission denied" }, { 71, "Disk not ready" }, { 74, "Can't rename with different drive" }, { 75, "Path/File access error" },
            { 76, "Path not found" }, { 91, "Object variable not set" }, { 92, "For loop not initialized" }, { 94, "Invalid use of Null" },
            { 322, "Can't create necessary temporary file" }, { 424, "Object required" }, { 429, "ActiveX component can't create object" },
            { 430, "Class doesn't support Automation" }, { 432, "File name or class name not found during Automation operation" },
            { 438, "Object doesn't support this property or method" }, { 440, "Automation error" }, { 445, "Object doesn't support this action" },
            { 446, "Object doesn't support named arguments" }, { 447, "Object doesn't support current locale setting" },
            { 448, "Named argument not found" }, { 449, "Argument not optional" }, { 450, "Wrong number of arguments or invalid property assignment" },
            { 451, "Object not a collection" }, { 453, "Specified DLL function not found" }, { 455, "Code resource lock error" },
            { 457, "This key is already associated with an element of this collection" }, { 458, "Variable uses an Automation type not supported in VBScript" },
            { 462, "The remote server machine does not exist or is unavailable" }, { 481, "Invalid picture" }, { 500, "Variable is undefined" },
            { 501, "Illegal assignment" }, { 502, "Object not safe for scripting" }, { 503, "Object not safe for initializing" },
            { 504, "Object not safe for creating" }, { 505, "Invalid or unqualified reference" }, { 506, "Class not defined" },
            { 507, "An exception occurred" }, { 5008, "Illegal assignment" }, { 5017, "Syntax error in regular expression" },
            { 5018, "Unexpected quantifier" }, { 5019, "Expected ']' in regular expression" }, { 5020, "Expected ')' in regular expression" },
            { 5021, "Invalid range in character set" },
        };

        private VBScriptErrorInfo(int number, string description, string source)
        {
            Number = number;
            Description = description;
            Source = source;
        }

        public int Number { get; }

        public string Description { get; }

        public string Source { get; }

        /// <summary>
        /// The standard description of a VBScript runtime error, or null if the number is not one
        /// </summary>
        public static string? StandardDescription(int number) => Descriptions.TryGetValue(number, out string? description) ? description : null;

        public static VBScriptErrorInfo From(Exception exception)
        {
            if (exception == null)
                throw new ArgumentNullException(nameof(exception));
            while ((((exception is TargetInvocationException) || (exception is AggregateException)) && (exception.InnerException != null))
                || ((exception is ArgumentException) && ((exception.InnerException is IndexOutOfRangeException) || (exception.InnerException is SpecificVBScriptException))))
            {
                // the array access wraps the error of the index (out of range: "Subscript out of range", not a number: "Type mismatch")
                exception = exception.InnerException!;
            }

            if (exception is CustomException custom)
            {
                // Err.Raise (and the errors of the built-in components): the description and source given, otherwise those of the VBScript error
                string? standard = StandardDescription(custom.ErrorNumber);
                string description = (custom.ErrDescription.Length > 0) ? custom.ErrDescription : (standard ?? "Unknown runtime error");
                string source = (custom.ErrSource.Length > 0) ? custom.ErrSource : ((standard != null) ? RuntimeErrorSource : "");
                return new VBScriptErrorInfo(custom.ErrorNumber, description, source);
            }
            if (exception is SpecificVBScriptException specific)
                return new VBScriptErrorInfo(specific.ErrorNumber, StandardDescription(specific.ErrorNumber) ?? specific.BasicDescription, RuntimeErrorSource);

            int number = NumberOf(exception);
            return new VBScriptErrorInfo(number, StandardDescription(number) ?? exception.Message, (StandardDescription(number) != null) ? RuntimeErrorSource : (exception.Source ?? ""));
        }

        /// <summary>
        /// The VBScript error that a .NET exception corresponds to (the HRESULT for others, eg. the errors of COM components)
        /// </summary>
        private static int NumberOf(Exception exception)
        {
            switch (exception)
            {
                case DivideByZeroException _: return 11;
                case OverflowException _: return 6;
                case InvalidCastException _: return 13;
                case FormatException _: return 13;
                case IndexOutOfRangeException _: return 9;
                case ArgumentOutOfRangeException _: return 9;
                case FileNotFoundException _: return 53;
                case DirectoryNotFoundException _: return 76;
                case UnauthorizedAccessException _: return 70;
                case OutOfMemoryException _: return 7;
                case MissingMemberException _: return 438;
                case NullReferenceException _: return 91;
                default: return exception.HResult;
            }
        }
    }
}
