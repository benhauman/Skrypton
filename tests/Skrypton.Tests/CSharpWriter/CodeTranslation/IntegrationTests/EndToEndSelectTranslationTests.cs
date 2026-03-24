using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndSelectTranslationTests : TestBase
    {
        /// <summary>
        /// This tests a fix made to select block translation - it was looking for token types based upon their content, rather than their type (so it was mistaking
        /// a StringToken whose content was a single comma characters as being an ArgumentSeparatorToken, if the type of the token is checked instead of its content
        /// then this sort of mistake will no longer occur)
        /// </summary>
        [TestMethod, MyFact]
        public void AllowSpecialCharactersToBeUsedAsStringsInSelectCases()
        {
            var source = @"
				Select Case x
					Case ""(""
						WScript.Echo ""Open""
					Case "")""
						WScript.Echo ""Close""
					Case "",""
						WScript.Echo ""Split""
				End Select";

            var expected = @"
				if (_.IF(_.EQ(_env.x, ""("")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Open"");
				}
				else if (_.IF(_.EQ(_env.x, "")"")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Close"");
				}
				else if (_.IF(_.EQ(_env.x, "","")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Split"");
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source, ["SKY101"]);
        }

		[TestMethod]
		public void SelectCaseWithStringTokens()
		{
            var source = @"
    Dim Size : Size = 0
	Suffix = "" B"" 
	Select Case Suffix 
		Case "" KB"" Size = Round(Size / 1024, 2) 
		Case "" MB""	Size = Round(Size / 1048576, 2) 
	End Select
";
            base.TestCSharpCodeTranslation(source, ["SKY101"]);
        }


        [TestMethod]
        public void XMultipleTokensOnTheCaseLine1() // from CT127_dialog_67
        {
            var sourceX = @"
SUB PriorityMatrix()
	Dim impact
	Dim urgency
	Dim impactText
	Dim urgencyText
	Dim priority
	Dim priorityText
	urgencyText = hlObj.GetValue(""IncidentAttribute.Urgency"",0,0,0,0)
	impactText = hlObj.GetValue(""IncidentAttribute.Impact"", 0,0,0,0)
	
	Select Case impactText
		Case ""ImpactSinglePerson"" impact = 1
		Case ""ImpactMultipleGroups"" impact = 2
		Case ""ImpactEntireOrganization"" impact = 3
		Case """" impact = 0
		Case Else impact = ComboBoxImpact.GetCurSel()
	End Select
	
	Select Case priority 
		Case 1 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityNormal""
		Case 2 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityMedium""
		Case 3 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityHigh""
		Case 4 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityUrgent""
		Case 5 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityCritical""
		Case Else hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""""
	End Select

END SUB
";
			var source = sourceX;
            source = @"
SUB PriorityMatrix()
	Dim impact
	Dim urgency
	Dim impactText
	Dim urgencyText
	Dim priority
	Dim priorityText
	impactText = hlObj.GetValue(""IncidentAttribute.Impact"", 0,0,0,0)
	
	Select Case impactText
		Case ""ImpactSinglePerson"" impact = 1
		Case ""ImpactMultipleGroups"" impact = 2
		Case ""ImpactEntireOrganization"" impact = 3
		Case """" impact = 0
		Case Else impact = ComboBoxImpact.GetCurSel()
	End Select

END SUB
";

            TestCSharpCodeTranslation(source, ["SKY102", "SKY107"]);
        }

        [TestMethod]
		[Ignore] // todo
        public void XMultipleTokensOnTheCaseLine2() // from CT127_dialog_67
        {
            var source = @"
SUB PriorityMatrix()
	Dim impact
	Dim urgency
	Dim impactText
	Dim urgencyText
	Dim priority
	Dim priorityText
	
	Select Case priority 
		Case 1 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityNormal""
		Case 2 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityMedium""
		Case 3 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityHigh""
		Case 4 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityUrgent""
		Case 5 hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""PriorityCritical""
		Case Else hlObj.SetValue ""CaseGeneral.Priority"" ,0,0,0,""""
	End Select

END SUB
";

            TestCSharpCodeTranslation(source, ["SKY102", "SKY107"]);
        }
    }
}
