using System;
using System.Collections;
using Skrypton.RuntimeSupport;
namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }
        protected override GlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) => new GlobalReferences(compatLayer, env);
        protected override void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences globalReferences)
        {
            var _env = env ?? throw new ArgumentNullException(nameof(env));
            var _outer = globalReferences ?? throw new ArgumentNullException(nameof(globalReferences));
        }
    }
    public sealed class GlobalReferences : GlobalReferencesBaseT<EnvironmentReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly GlobalReferences _outer;
        private readonly EnvironmentReferences _env;
        public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) : base(compatLayer, env)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = this;
        }
        public void Button1_Click()
        {
            object URL = null;
            object wshShell = null;
            object oExec = null;
            object Processes = null; /* Undeclared in source */
            object intProcessId = null; /* Undeclared in source */
            object Process = null; /* Undeclared in source */

            URL = _.VAL(_.CALLm1v5(this, _.NnO(_env.hlObj, "hlObj"), "GetValue", "vRealize.LansweeperURL", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            wshShell = _.OBJ(_.CREATEOBJECT("WScript.Shell"));
            _.CALLm1argp(this, _.NnO(wshShell, "wshShell"), "run", _.ARGS.Ref(URL, v => { URL = v; }));

            Processes = _.OBJ(_.CALLm1v1(this, _.NnO(_.GETOBJECT("winmgmts:"), "(GetObject result)"), "InstancesOf", "Win32_Process"));

            intProcessId = "";
            var enumerationContent = _.ENUMERABLE(Processes).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                Process = enumerationContent.Current;
                if (_.IF(_.EQ(_.NullableNUM(_.STRCOMP(_.CALLm1v0(this, _.NnO(Process, "Process"), "Name"), "iexplore.exe", VBScriptConstants.vbTextCompare)), (Int16)0)))
                {
                    intProcessId = _.VAL(_.CALLm1v0(this, _.NnO(Process, "Process"), "ProcessId"));
                    break;
                }
            }

            if (_.IF(_.GT(_.NullableNUM(_.LEN(intProcessId)), (Int16)0)))
            {
                var with = _.OBJ(_.CREATEOBJECT("WScript.Shell"));
                _.CALLm1argp(this, _.NnO(with, "with"), "AppActivate", _.ARGS.Ref(intProcessId, v2 => { intProcessId = v2; }));

            }
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object Button1_Click { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlObj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
