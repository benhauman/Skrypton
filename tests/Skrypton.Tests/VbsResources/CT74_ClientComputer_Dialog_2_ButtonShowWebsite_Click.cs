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

        public void ButtonShowWebsite_Click()
        {
            object sServer = null;
            object sConn = null;
            object oConn = null;
            object sDatabaseName = null;
            object sUser = null;
            object sPassword = null;
            const Int16 adCmdStoredProc = (Int16)4;
            const Int16 adInteger = (Int16)3;
            const Int16 adVarWChar = (Int16)202;
            const Int16 adParamInput = (Int16)1;
            const Int16 adParamOutput = (Int16)2;
            const Int16 adParamReturnValue = (Int16)4;
            object parmname = null;
            object parmval = null;
            object FirstCharName = null;
            object xvIdentifier = null;
            object rewritten_group = null;
            object adoSQLCmdParam = null; /* Undeclared in source */
            object adoSQLCmdParam2 = null; /* Undeclared in source */

            //If hlObj.HasContent("PersonBilling.CostCenter_CA",0,0) = 0 Then
            //	model.MsgBox "Bitte zuerst eine Kostenstelle erfassen!"
            //	model.CurrentCommand.abort"OnSave"
            //End if

            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, _env.hlObj, "IsNew")), (Int16)1)))
            {

                sDatabaseName = "HLData";
                sServer = "MSSQLB";
                sUser = "helplinedata";
                sPassword = "helplinedata";
                sConn = _.CONCAT("provider=sqloledb;data source=", sServer, ";initial catalog=", sDatabaseName);
                oConn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.CALLm1argp(this, oConn, "Open", _.ARGS.Ref(sConn, v => { sConn = v; }).Ref(sUser, v2 => { sUser = v2; }).Ref(sPassword, v3 => { sPassword = v3; }));

                FirstCharName = _.VAL(_.LEFT(_.CALLm1argp(this, _env.hlObj, "GetValue", _.ARGS.Val("PersonGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), (Int16)1));

                //SB Code ermitteln
                parmname = "runScript";
                adoSQLCmdParam = _.OBJ(_.CREATEOBJECT("ADODB.Command"));
                var with = _.OBJ(adoSQLCmdParam);
                _.SET(_.OBJ(oConn), this, with, "ActiveConnection");
                _.SET("CreateNewSBCode", this, with, "CommandText");
                _.SET(_.VAL(adCmdStoredProc), this, with, "CommandType");
                _.CALLm2argp(this, with, "Parameters", "Append", _.ARGS.Val(_.CALLm1v3(this, with, "CreateParameter", "RETURN_VALUE", adInteger, adParamReturnValue)));
                _.CALLm2argp(this, with, "Parameters", "Append", _.ARGS.Val(_.CALLm1argp(this, with, "CreateParameter", _.ARGS.Val("@FirstCharName").Val(adVarWChar).Val(adParamInput).Val((Int16)1).Ref(FirstCharName, v4 => { FirstCharName = v4; }))));
                _.CALLm2argp(this, with, "Parameters", "Append", _.ARGS.Val(_.CALLm1v4(this, with, "CreateParameter", "@NewSBCode", adVarWChar, adParamOutput, (Int16)10)));
                _.CALLm1v0(this, with, "Execute");
                parmval = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, with, "Parameters", (Int16)2), "Value"));

                _.CALLm1argp(this, _env.hlObj, "SetValue", _.ARGS.Val("PersonInformation.SBCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(parmval, v5 => { parmval = v5; }));

                rewritten_group = _.VAL(_.CALLm1argp(this, _env.hlObj, "GetValue", _.ARGS.Val("PersonGeneral.Group").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

                if (_.IF(_.OR(_.EQ(_.NullableSTR(rewritten_group), "GroupMainova"), _.EQ(_.NullableSTR(rewritten_group), "GroupHolding"))))
                {
                    xvIdentifier = "X";
                }
                else
                {
                    xvIdentifier = "V";
                }

                //X/V Personalnummer ermitteln
                adoSQLCmdParam2 = _.OBJ(_.CREATEOBJECT("ADODB.Command"));
                var with2 = _.OBJ(adoSQLCmdParam2);
                _.SET(_.OBJ(oConn), this, with2, "ActiveConnection");
                _.SET("CreateNewPersonalID", this, with2, "CommandText");
                _.SET(_.VAL(adCmdStoredProc), this, with2, "CommandType");
                _.CALLm2argp(this, with2, "Parameters", "Append", _.ARGS.Val(_.CALLm1v3(this, with2, "CreateParameter", "RETURN_VALUE", adInteger, adParamReturnValue)));
                _.CALLm2argp(this, with2, "Parameters", "Append", _.ARGS.Val(_.CALLm1argp(this, with2, "CreateParameter", _.ARGS.Val("@TypeCode").Val(adVarWChar).Val(adParamInput).Val((Int16)1).Ref(xvIdentifier, v6 => { xvIdentifier = v6; }))));
                _.CALLm2argp(this, with2, "Parameters", "Append", _.ARGS.Val(_.CALLm1v4(this, with2, "CreateParameter", "@NewPersonalID", adVarWChar, adParamOutput, (Int16)10)));
                _.CALLm1v0(this, with2, "Execute");
                parmval = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, with2, "Parameters", (Int16)2), "Value"));

                _.CALLm1argp(this, _env.hlObj, "SetValue", _.ARGS.Val("PersonGeneral.PersonalID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(parmval, v7 => { parmval = v7; }));

                _.CALLm1v0(this, oConn, "Close");
                oConn = VBScriptConstants.Nothing;

            }
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object ButtonShowWebsite_Click { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlObj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
