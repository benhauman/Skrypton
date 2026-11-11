#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in ADODB (Connection, Recordset, Command, Stream) on the host's IHostDatabaseConnectionFactoryHostService
    /// </summary>
    [TestClass]
    public sealed class ScriptonADODBTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        private static DataTable People()
        {
            var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add(1, "Ann");
            table.Rows.Add(2, DBNull.Value);
            table.Rows.Add(3, "Cid");
            return table;
        }

        private List<object?> Execute(string vbs, TestDatabaseConnectionFactoryHostService? database)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (database != null)
                        s.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => database);
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        private const string OpenConnection =
            "Dim cn : Set cn = CreateObject(\"ADODB.Connection\")\r\n" +
            "cn.Open \"Provider=SQLOLEDB;Data Source=test\"\r\n";

        [TestMethod]
        public void ExecuteAndNavigateARecordset()
        {
            var database = new TestDatabaseConnectionFactoryHostService(sql => sql.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) ? People() : null);
            var values = Execute(
                OpenConnection +
                "Dim rs : Set rs = cn.Execute(\"SELECT id, name FROM people\")\r\n" +
                "collector.Report rs.RecordCount\r\n" +
                "Dim s\r\n" +
                "Do Until rs.EOF\r\n" +
                "  s = s & rs(\"id\") & \"=\" & IsNull(rs.Fields(\"name\").Value) & \";\"\r\n" +
                "  rs.MoveNext\r\n" +
                "Loop\r\n" +
                "collector.Report s\r\n" +
                "rs.MoveFirst\r\n" +
                "collector.Report rs.Fields(1).Name & \"|\" & rs(\"name\") & \"|\" & rs.Fields.Count\r\n" +
                "rs.Close\r\n" +
                "collector.Report rs.State & \"|\" & cn.State\r\n" +
                "cn.Close\r\n" +
                "collector.Report cn.State\r\n",
                database);
            CollectionAssert.AreEqual(new object?[] { 3, "1=False;2=True;3=False;", "name|Ann|2", "0|1", 0 }, values.ToArray());
            Assert.AreEqual("Provider=SQLOLEDB;Data Source=test", database.ConnectionStrings.Single());
        }

        [TestMethod]
        public void RecordsetOpenWithAConnection()
        {
            var database = new TestDatabaseConnectionFactoryHostService(_ => People());
            var values = Execute(
                OpenConnection +
                "Dim rs : Set rs = CreateObject(\"ADODB.Recordset\")\r\n" +
                "rs.Open \"SELECT * FROM people\", cn\r\n" +
                "rs.MoveLast\r\n" +
                "collector.Report rs(\"name\").Value\r\n",
                database);
            Assert.AreEqual("Cid", values[0]);
        }

        [TestMethod]
        public void CommandWithParameters()
        {
            var database = new TestDatabaseConnectionFactoryHostService();
            Execute(
                OpenConnection +
                "Dim cmd : Set cmd = CreateObject(\"ADODB.Command\")\r\n" +
                "Set cmd.ActiveConnection = cn\r\n" +
                "cmd.CommandText = \"UPDATE people SET name = ? WHERE id = ?\"\r\n" +
                "cmd.Parameters.Append cmd.CreateParameter(\"name\", 200, 1, 50, \"Bob\")\r\n" +
                "cmd.Parameters.Append cmd.CreateParameter(\"id\", 3, 1, 0, 2)\r\n" +
                "cmd.Execute\r\n",
                database);
            TestDbCommand command = database.ExecutedCommands.Single();
            Assert.AreEqual("UPDATE people SET name = ? WHERE id = ?", command.CommandText);
            CollectionAssert.AreEqual(new object[] { "name=Bob:AnsiString", "id=2:Int32" },
                command.ExecutedParameters.Select(p => $"{p.ParameterName}={p.Value}:{p.DbType}").ToArray());
        }

        [TestMethod]
        public void TransactionsAndStatementsWithoutRows()
        {
            var database = new TestDatabaseConnectionFactoryHostService();
            var values = Execute(
                OpenConnection +
                "cn.BeginTrans\r\n" +
                "Dim rs : Set rs = cn.Execute(\"DELETE FROM people\")\r\n" +
                "collector.Report rs.State\r\n" +
                "cn.CommitTrans\r\n",
                database);
            Assert.AreEqual(0, values[0], "a statement that returns no rows gives a closed Recordset");
        }

        [TestMethod]
        public void WithoutAHostServiceOpenRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute(OpenConnection, database: null));
            StringAssert.Contains(exception.ToString(), nameof(IHostDatabaseConnectionFactoryHostService));
        }

        [TestMethod]
        public void StreamTextRoundTrip()
        {
            var values = Execute(
                "Dim st : Set st = CreateObject(\"ADODB.Stream\")\r\n" +
                "st.Type = 2 : st.Charset = \"utf-8\"\r\n" +
                "st.Open\r\n" +
                "st.WriteText \"Hello\"\r\n" +
                "st.Position = 0\r\n" +
                "collector.Report st.ReadText()\r\n" +
                "collector.Report st.Size\r\n" +
                "st.Close\r\n",
                database: null);
            // Size includes the byte order mark of UTF-8 (verified with Windows ADO: "Hello" is 8 bytes)
            CollectionAssert.AreEqual(new object?[] { "Hello", 8 }, values.ToArray());
        }

        [TestMethod]
        [DataRow("utf-8", 5)]
        [DataRow("Unicode", 6)]
        [DataRow("iso-8859-1", 2)]
        public void StreamTextStartsWithTheByteOrderMarkOfUtf8AndUnicode(string charset, int expectedSize)
        {
            // verified with Windows ADO: "Hi" is 5 bytes in UTF-8 (EF BB BF 48 69), 6 in Unicode (FF FE 48 00 69 00), 2 in ISO-8859-1
            var values = Execute(
                "Dim st : Set st = CreateObject(\"ADODB.Stream\")\r\n" +
                "st.Type = 2 : st.Charset = \"" + charset + "\" : st.Open\r\n" +
                "st.WriteText \"Hi\"\r\n" +
                "collector.Report st.Size\r\n" +
                "st.Position = 0\r\n" +
                "collector.Report st.ReadText()\r\n",
                database: null);
            CollectionAssert.AreEqual(new object?[] { expectedSize, "Hi" }, values.ToArray());
        }
    }
}
