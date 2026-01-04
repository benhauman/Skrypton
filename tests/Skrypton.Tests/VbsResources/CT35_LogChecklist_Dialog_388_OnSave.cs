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

            //Check if invalid characters are in any of the url textboxes
            _env.dict = _.OBJ(_.CREATEOBJECT("Scripting.Dictionary"));
            _env.objregex = _.OBJ(_.CREATEOBJECT("VBScript.RegExp"));
            _.SET(true, this, _env.objregex, "Global");
            _.SET("[^A-Z0-9][^\\:][^\\/][^\\.][^\\S][^\\?][^\\€][^\\@]", this, _env.objregex, "Pattern");

            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 1 URL").Val(_.CALL(this, _env.textboxchecklist1url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 2 URL").Val(_.CALL(this, _env.textboxchecklist2url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 3 URL").Val(_.CALL(this, _env.textboxchecklist3url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 4 URL").Val(_.CALL(this, _env.textboxchecklist4url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 5 URL").Val(_.CALL(this, _env.textboxchecklist5url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 6 URL").Val(_.CALL(this, _env.textboxchecklist6url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 7 URL").Val(_.CALL(this, _env.textboxchecklist7url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 8 URL").Val(_.CALL(this, _env.textboxchecklist8url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 9 URL").Val(_.CALL(this, _env.textboxchecklist9url, "Text")));
            _.CALL(this, _env.dict, "Add", _.ARGS.Val("Checkliste 10 URL").Val(_.CALL(this, _env.textboxchecklist10url, "Text")));

            var enumerationContent = _.ENUMERABLE(_env.dict).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                _outer.element = enumerationContent.Current;
                if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.dict, _.ARGS.Ref(_outer.element, v => { _outer.element = v; }))), "")))
                {
                    _env.match = _.OBJ(_.CALL(this, _env.objregex, "execute", _.ARGS.RefIfArray(_env.dict, _.ARGS.Ref(_outer.element, v2 => { _outer.element = v2; }))));
                    if (_.IF(_.GT(_.NullableNUM(_.CALL(this, _env.match, "Count")), (Int16)0)))
                    {
                        _outer.errmsg = _.VAL(_.CALL(this, _env.model, "Translate", _.ARGS.Val("#ERR_Checklists_InvalidChars")));
                        _outer.errmsg = _.REPLACE(_outer.errmsg, "{0}", _outer.element);
                        _.CALL(this, _env.model, "MsgBox", _.ARGS.Ref(_outer.errmsg, v5 => { _outer.errmsg = v5; }));
                        _.CALL(this, _env.model, "CurrentCommand", "Abort", _.ARGS.Val("OnSave"));
                    }
                }
            }
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
            element = null;
            errmsg = null;
        }

        internal object element { get; set; }
        internal object errmsg { get; set; }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object dict { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object match { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object objregex { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist10url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist1url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist2url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist3url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist4url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist5url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist6url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist7url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist8url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist9url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}