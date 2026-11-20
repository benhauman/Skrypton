using System;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// A script engine that compiles a script eagerly, once, and then only binds host objects and calls procedures (issue #190) - for a host that
    /// runs one large script version for many requests. Unlike IScriptControl (the MSScriptControl replacement) there is no AddCode, AddObject,
    /// ExecuteStatement, Eval, Modules or Reset: a new code version means a new engine. The phases are strict, a call out of order raises an
    /// InvalidOperationException:
    /// <list type="number">
    /// <item>DeclareObject - the host object names and contract types, without instances</item>
    /// <item>SetCode - the whole script, once</item>
    /// <item>CompileCode - translate and compile (once per version, process-wide), runs nothing, seals the engine</item>
    /// <item>BindObject - the instances of the declared objects, per run</item>
    /// <item>RunGlobalCode / CallProcedure - the global code runs once per engine, the module state is kept between calls</item>
    /// </list>
    /// An engine is not thread-safe (one per request); the compiled code handle is.
    /// </summary>
    public interface IEagerScriptEngine : IDisposable
    {
        /// <summary>
        /// Phase 1: declare a host object by name and contract type (no instance). typeof(object) (or a COM interface without a managed type)
        /// means late-bound. With addMembers the public methods of the type are callable without the object name (as AddObject's AddMembers).
        /// </summary>
        void DeclareObject(string name, Type type, bool addMembers = false);

        /// <summary>
        /// Phase 2: the whole script, once
        /// </summary>
        void SetCode(string code);

        /// <summary>
        /// Phase 3: translate and compile the script through the process-wide cache (an unchanged version is not compiled again, parallel callers
        /// of a new version wait for one compilation, a failed compilation is not cached) and seal the engine. Runs no script code and needs no
        /// host object instances. Compile errors are raised here (and described by Error). The returned handle creates further engines of this
        /// version; dispose it to retire the version.
        /// </summary>
        CompiledCodeHandle CompileCode();

        /// <summary>
        /// Phase 4: the instance of a declared host object for this run (before the first RunGlobalCode / CallProcedure) - it must be an instance
        /// of the declared type
        /// </summary>
        void BindObject(string name, object instance);

        /// <summary>
        /// The Subs and Functions of the compiled script (from phase 4 on)
        /// </summary>
        Procedures Procedures { get; }

        bool ContainsProcedure(string name);

        /// <summary>
        /// Phase 5: run the global (outer scope) code of the script - once per engine, every declared object must be bound. A second call raises
        /// an InvalidOperationException; procedures called afterwards see the global variables it set.
        /// </summary>
        void RunGlobalCode();

        /// <summary>
        /// Phase 5: call a Sub or Function of the script (the global code runs first if RunGlobalCode was not called). A ByRef parameter that the
        /// procedure changes updates its element of args. No parse, no translation, no compile.
        /// </summary>
        object? CallProcedure(string name, params object?[] args);

        /// <summary>
        /// The error raised by the most recent CompileCode, RunGlobalCode or CallProcedure (Number is zero if there was none), with the VBScript
        /// line and column
        /// </summary>
#pragma warning disable CA1716 // Identifiers should not match keywords (as IScriptControl.Error)
        Error Error { get; }
#pragma warning restore CA1716
    }

    /// <summary>
    /// Which engine a ScriptEngineConfiguration is for (issue #190)
    /// </summary>
    public enum ScriptEngineMode
    {
        /// <summary>
        /// IScriptControl / ScriptControlClass, the MSScriptControl replacement
        /// </summary>
        ScriptControl = 0,

        /// <summary>
        /// IEagerScriptEngine / EagerScriptEngine: declare, set code, compile once, then bind and call procedures
        /// </summary>
        EagerScript = 1,
    }

    /// <summary>
    /// A ScriptControlConfiguration with the engine mode (issue #190) - EagerScriptEngine requires ScriptEngineMode.EagerScript, ScriptControlClass
    /// refuses it
    /// </summary>
    public class ScriptEngineConfiguration : ScriptControlConfiguration
    {
        public ScriptEngineConfiguration(bool tempEnabled, string? tempDirectoryPath, bool enabledLoadFromDisk, string[] translationSuppression, string[] noWarn, bool treatWarningsAsErrors = true, bool undeclaredVariablesAreErrors = true, ScriptEngineMode mode = ScriptEngineMode.ScriptControl)
            : base(tempEnabled, tempDirectoryPath, enabledLoadFromDisk, translationSuppression, noWarn, treatWarningsAsErrors, undeclaredVariablesAreErrors)
        {
            if (!Enum.IsDefined(typeof(ScriptEngineMode), mode))
                throw new ArgumentOutOfRangeException(nameof(mode));
            Mode = mode;
        }

        public ScriptEngineMode Mode { get; }
    }
}
