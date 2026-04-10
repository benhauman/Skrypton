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
            //Check if invalid characters are in any of the url textboxes
            _env.dict = _.OBJ(_.CREATEOBJECT("Scripting.Dictionary"));
            _env.objRegEx = _.OBJ(_.CREATEOBJECT("VBScript.RegExp"));
            _.SETm1a0(this, _env.objRegEx ?? throw new InvalidOperationException("Reference not set:objRegEx"), "Global", true);
            _.SETm1a0(this, _env.objRegEx ?? throw new InvalidOperationException("Reference not set:objRegEx"), "Pattern", "[^A-Z0-9][^\\:][^\\/][^\\.][^\\S][^\\?][^\\€][^\\@]");

            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 1 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist1URL, "TextBoxChecklist1URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 2 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist2URL, "TextBoxChecklist2URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 3 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist3URL, "TextBoxChecklist3URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 4 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist4URL, "TextBoxChecklist4URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 5 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist5URL, "TextBoxChecklist5URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 6 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist6URL, "TextBoxChecklist6URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 7 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist7URL, "TextBoxChecklist7URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 8 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist8URL, "TextBoxChecklist8URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 9 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist9URL, "TextBoxChecklist9URL"), "Text"));
            _.CALLm1v2(this, _.NnO(_env.dict, "dict"), "Add", "Checkliste 10 URL", _.CALLm1v0(this, _.NnO(_env.TextBoxChecklist10URL, "TextBoxChecklist10URL"), "Text"));

            var enumerationContent = _.ENUMERABLE(_env.dict).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                _outer.element = enumerationContent.Current;
                if (_.IF(_.NOTEQ(_.NullableSTR(_.CALLm0argp(this, _.NnO(_env.dict, "dict"), _.ARGS.Ref(_outer.element, v => { _outer.element = v; }))), "")))
                {
                    _env.match = _.OBJ(_.CALLm1argp(this, _.NnO(_env.objRegEx, "objRegEx"), "execute", _.ARGS.RefIfArray(_env.dict, _.ARGS.Ref(_outer.element, v2 => { _outer.element = v2; }))));
                    if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, _.NnO(_env.match, "match"), "Count")), (Int16)0)))
                    {
                        _outer.errMsg = _.VAL(_.CALLm1v1(this, _.NnO(_env.model, "model"), "Translate", "#ERR_Checklists_InvalidChars"));
                        _outer.errMsg = _.REPLACE(_outer.errMsg, "{0}", _outer.element);
                        _.CALLm1argp(this, _.NnO(_env.model, "model"), "MsgBox", _.ARGS.Ref(_outer.errMsg, v3 => { _outer.errMsg = v3; }));
                        _.CALLm2v1(this, _.NnO(_env.model, "model"), "CurrentCommand", "Abort", "OnSave");
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
            errMsg = null;
        }
        internal object element { get; set; }
        internal object errMsg { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object dict { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object match { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object objRegEx { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist10URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist1URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist2URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist3URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist4URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist5URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist6URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist7URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist8URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist9URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
