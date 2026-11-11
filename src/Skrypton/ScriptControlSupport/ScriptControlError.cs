using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Skrypton.RuntimeSupport;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// IScriptControl.Error - describes the error raised by the most recent ExecuteStatement call, like MSScriptControl's Error object (see issue #47)
    /// </summary>
    internal sealed class ScriptControlError : Error
    {
        internal const string RuntimeErrorSource = "Microsoft VBScript runtime error";
        internal const string CompilationErrorSource = "Microsoft VBScript compilation error";

        public int Number { get; private set; }
        public string Source { get; private set; } = "";
        public string Description { get; private set; } = "";
        public string? HelpFile { get; private set; }
        public int HelpContext { get; private set; }
        public string Text { get; private set; } = "";
        public int Line { get; private set; }
        public int Column { get; private set; }

        public void Clear()
        {
            Number = 0;
            Source = "";
            Description = "";
            HelpFile = null;
            HelpContext = 0;
            Text = "";
            Line = 0;
            Column = 0;
        }

        /// <summary>
        /// Populate from an exception raised while translating, compiling or running a script. If programAssemblyIfCompiled is null then the
        /// error happened before the program could be run (so it is a compilation error). The source line map (1-based C# line to VBScript source
        /// line index) is used to report the VBScript line of a runtime error.
        /// </summary>
        internal void SetFrom(Exception exception, Assembly? programAssemblyIfCompiled, IReadOnlyDictionary<int, int>? sourceLineMap, string? scriptContent, string? csCode = null)
        {
            if (exception == null)
                throw new ArgumentNullException(nameof(exception));

            Exception error = Unwrap(exception);
            bool isCompilationError = programAssemblyIfCompiled == null;
            SpecificVBScriptException? specificError = error as SpecificVBScriptException;

            // runtime errors: number and source as Windows VBScript reports them (Err.Raise keeps its own source, issue #85)
            Skrypton.RuntimeSupport.Exceptions.VBScriptErrorInfo? runtimeError = isCompilationError ? null : Skrypton.RuntimeSupport.Exceptions.VBScriptErrorInfo.From(error);
            Number = (specificError != null) ? specificError.ErrorNumber : (isCompilationError ? 1002 : runtimeError!.Number); // 1002: Syntax error
            Source = isCompilationError ? CompilationErrorSource : ((runtimeError!.Source.Length > 0) ? runtimeError.Source : RuntimeErrorSource);
            Description = GetDescription(error);
            HelpFile = null;
            HelpContext = 0;
            Column = 0;
            int csLine = 0;
            Line = isCompilationError ? 0 : TryGetSourceLine(error, programAssemblyIfCompiled!, sourceLineMap, out csLine);
            if (error is Skrypton.RuntimeSupport.Exceptions.VariableIsUndefinedException undefined)
            {
                // found by the translation, but a runtime error of VBScript (with the line of the variable)
                Source = RuntimeErrorSource;
                Line = undefined.ScriptLine;
            }
            // the column of the statement on the line (0-based, as MSScriptControl's Error.Column - issue #87)
            Column = (Line > 0) ? ScriptStatementColumns.Find(GetSourceLine(scriptContent, Line), GetSourceLine(csCode, csLine), (error as Skrypton.RuntimeSupport.Exceptions.VariableIsUndefinedException)?.HostTarget) : 0;
            Text = (Line > 0) ? GetSourceLineText(scriptContent, Line) : "";
        }

        private static Exception Unwrap(Exception exception)
        {
            while ((exception is TargetInvocationException || exception is AggregateException) && (exception.InnerException != null))
                exception = exception.InnerException;
            return exception;
        }

        private static string GetDescription(Exception error)
        {
            // Err.Raise and .NET exceptions: as the Err object has them; VBScript errors: with what VBScript names when it reports them to the host
            // ("Object required: 'name'", not in Err.Description - issue #84)
            if ((error is Skrypton.RuntimeSupport.Exceptions.CustomException) || !(error is SpecificVBScriptException))
                return Skrypton.RuntimeSupport.Exceptions.VBScriptErrorInfo.From(error).Description;
            if (((SpecificVBScriptException)error).HostTarget != null)
                return ((SpecificVBScriptException)error).HostDescription;

            // SpecificVBScriptException messages are "<basic description>: <additional information>" and the basic description is what
            // MSScriptControl reports (eg. "Division by zero")
            if (error is SpecificVBScriptException)
            {
                int separatorIndex = error.Message.IndexOf(": ", StringComparison.Ordinal);
                return (separatorIndex > 0) ? error.Message.Substring(0, separatorIndex) : error.Message;
            }
            return error.Message;
        }

        /// <summary>
        /// Find the innermost stack frame in the compiled script program and map its C# line back to the 1-based VBScript line (0 if unknown)
        /// </summary>
        private static int TryGetSourceLine(Exception error, Assembly programAssembly, IReadOnlyDictionary<int, int>? sourceLineMap, out int csLineOfError)
        {
            csLineOfError = 0;
            if ((sourceLineMap == null) || (sourceLineMap.Count == 0))
                return 0;

            StackFrame[] frames = new StackTrace(error, fNeedFileInfo: true).GetFrames() ?? Array.Empty<StackFrame>();
            foreach (StackFrame frame in frames)
            {
                MethodBase? method = frame.GetMethod();
                if ((method?.DeclaringType == null) || (method.DeclaringType.Assembly != programAssembly))
                    continue;
                int csLine = frame.GetFileLineNumber();
                if (csLine <= 0)
                    continue;

                // The statement that this C# line belongs to started on the closest mapped line at or before it
                int[] candidates = sourceLineMap.Keys.Where(l => l <= csLine).ToArray();
                if (candidates.Length == 0)
                    continue;
                csLineOfError = csLine;
                return sourceLineMap[candidates.Max()];
            }
            return 0;
        }

        private static string GetSourceLine(string? content, int line)
        {
            if (string.IsNullOrEmpty(content))
                return "";
            string[] lines = content!.Replace("\r\n", "\n").Split('\n');
            return ((line >= 1) && (line <= lines.Length)) ? lines[line - 1] : "";
        }

        private static string GetSourceLineText(string? scriptContent, int line)
        {
            if (string.IsNullOrEmpty(scriptContent))
                return "";
            string[] lines = scriptContent!.Replace("\r\n", "\n").Split('\n');
            return ((line >= 1) && (line <= lines.Length)) ? lines[line - 1].Trim() : "";
        }
    }
}
