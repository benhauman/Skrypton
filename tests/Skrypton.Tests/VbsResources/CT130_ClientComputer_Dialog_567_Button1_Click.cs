using System;
using System.Collections;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Compat;

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

            URL = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("vRealize.LansweeperURL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            wshShell = _.OBJ(_.CREATEOBJECT("WScript.Shell"));
            _.CALL(this, wshShell, "run", _.ARGS.Ref(URL, v => { URL = v; }));

            Processes = _.OBJ(_.CALL(this, _.GETOBJECT("winmgmts:"), "InstancesOf", _.ARGS.Val("Win32_Process")));

            intProcessId = "";
            var enumerationContent = _.ENUMERABLE(Processes).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                Process = enumerationContent.Current;
                if (_.IF(_.EQ(_.NullableNUM(_.STRCOMP(_.CALL(this, Process, "Name"), "iexplore.exe", VBScriptConstants.vbTextCompare)), (Int16)0)))
                {
                    intProcessId = _.VAL(_.CALL(this, Process, "ProcessId"));
                    break;
                }
            }

            if (_.IF(_.GT(_.NullableNUM(_.LEN(intProcessId)), (Int16)0)))
            {
                var with = _.OBJ(_.CREATEOBJECT("WScript.Shell"));
                _.CALL(this, with, "AppActivate", _.ARGS.Ref(intProcessId, v2 => { intProcessId = v2; }));

            }
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object Button1_Click { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlobj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
