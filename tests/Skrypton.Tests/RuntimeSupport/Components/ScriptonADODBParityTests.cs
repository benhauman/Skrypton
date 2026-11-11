#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The expected lines were produced by Windows cscript with real ADO (SQLOLEDB) against SQL Server: an ADODB Field passed as an argument is the
    /// Field object (not its value), and VarType follows an object's parameter-less default property
    /// </summary>
    [TestClass]
    public sealed class ScriptonADODBParityTests : TestBase
    {
        public sealed class C { public List<object?> V { get; } = new List<object?>(); public void Report(object? o) => V.Add(o); }

        [TestMethod]
        public void SameResultsAsWindowsVBScriptWithRealAdo()
        {
            var table = new DataTable();
            table.Columns.Add("name", typeof(string));
            table.Columns.Add("empty", typeof(string));
            table.Rows.Add("Ann", DBNull.Value);
            var db = new TestDatabaseConnectionFactoryHostService(_ => table);
            var c = new C();
            IScriptControl sc = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => db))), CreateScriptControlConfiguration(false, [], []));
            sc.AddObject("collector", c, false);
            sc.ExecuteStatement(
                "Dim cn, rs\r\n" +
                "Set cn = CreateObject(\"ADODB.Connection\")\r\n" +
                "cn.Open \"x\"\r\n" +
                "Set rs = cn.Execute(\"SELECT 'Ann' AS name\")\r\n" +
                "Sub ReportByVal(ByVal v)\r\n  collector.Report \"ByVal: TypeName=\" & TypeName(v) & \" VarType=\" & VarType(v) & \" IsObject=\" & IsObject(v)\r\nEnd Sub\r\n" +
                "Sub ReportByRef(v)\r\n  collector.Report \"ByRef: TypeName=\" & TypeName(v) & \" VarType=\" & VarType(v) & \" IsObject=\" & IsObject(v)\r\nEnd Sub\r\n" +
                "ReportByVal rs(\"name\")\r\n" +
                "ReportByRef rs(\"name\")\r\n" +
                "ReportByVal rs.Fields(\"name\")\r\n" +
                "collector.Report \"TypeName(rs(\"\"name\"\"))=\" & TypeName(rs(\"name\"))\r\n" +
                "Dim d : Set d = CreateObject(\"Scripting.Dictionary\")\r\n" +
                "d.Add \"k\", rs(\"name\")\r\n" +
                "collector.Report \"Dictionary item: TypeName=\" & TypeName(d(\"k\"))\r\n" +
                "collector.Report \"Concat: \" & rs(\"name\")\r\n" +
                "collector.Report \"Null field: IsNull=\" & IsNull(rs(\"empty\")) & \" TypeName(rs(\"\"empty\"\").Value)=\" & TypeName(rs(\"empty\").Value)\r\n" +
                "collector.Report \"Dictionary=\" & VarType(CreateObject(\"Scripting.Dictionary\"))\r\n" +
                "collector.Report \"Nothing=\" & VarType(Nothing)\r\n" +
                "collector.Report \"Connection=\" & VarType(CreateObject(\"ADODB.Connection\"))\r\n");
            CollectionAssert.AreEqual(new object[]
            {
                "ByVal: TypeName=Field VarType=8 IsObject=True",
                "ByRef: TypeName=Field VarType=8 IsObject=True",
                "ByVal: TypeName=Field VarType=8 IsObject=True",
                "TypeName(rs(\"name\"))=Field",
                "Dictionary item: TypeName=Field",
                "Concat: Ann",
                "Null field: IsNull=True TypeName(rs(\"empty\").Value)=Null",
                "Dictionary=9",
                "Nothing=9",
                "Connection=8",
            }, c.V.ToArray());
        }
    }
}
