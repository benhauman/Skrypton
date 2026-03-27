using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndClassTranslationTests : TestBase
    {
        /// <summary>
        /// When the tokens with the content "Property" had to be classified as a MayBeKeywordOrNameToken instead of a straight KeyWordToken, some logic had to
        /// be changed in the class parsing to account for it - this test exercises that work
        /// </summary>
        [TestMethod, MyFact]
        public void EndProperty()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
					END PROPERTY
				END CLASS
			";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}

					[TranslatedProperty(""Name"")]
					public object Name()
					{
						return null;
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected,
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// If a class has a Class_Terminate method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
        /// implement IDisposable so that it's possible to instantiate it and tidy it up to simulate the way in which the deterministic VBScript interpreter
        /// calls Class_Terminate (as soon as it leaves scope, rather than when a garbage collector wants to deal with it). This isn't currently taken
        /// advantage of in the generated code (as of 2014-12-15) but it might be in the future.
        /// </summary>
        [TestMethod, MyFact]
        public void ClassTerminateResultsInDisposableTranslatedClass()
        {
            var source = @"
				CLASS C1
					PUBLIC SUB ClAsS_TeRmInAtE
						WScript.Echo ""Gone!""
					END SUB
				END CLASS
			";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1 : IDisposable
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					private bool _disposed;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
						_disposed = false;
					}

					~C1()
					{
						try { Dispose(false); }
						catch(Exception e)
						{
							try { _.SETERROR(e); } catch { }
						}
					}

					void IDisposable.Dispose()
					{
						Dispose(true);
						GC.SuppressFinalize(this);
					}

					private void Dispose(bool disposing)
					{
						if (_disposed)
							return;
						if (disposing)
							Class_Terminate();
						_disposed = true;
					}

					public void Class_Terminate()
					{
						_.CALLm1v1(this, _env.WScript, ""Echo"", ""Gone!"");
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected,//.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //sourece//WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// If a class has a Class_Initialize method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
        /// call this method in the constructor in the generated class. For strict compatibility with VBScript, any error is ignored and, while it will terminate
        /// execution of Class_Initialize, it will not prevent the calling code from continuing.
        /// </summary>
        [TestMethod, MyFact]
        public void ClassInitializeResultsInConstructorCall()
        {
            var source = @"
				CLASS C1
					PUBLIC SUB ClAsS_InItIaLiZe
						WScript.Echo ""Here!""
					END SUB
				END CLASS
			";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
						try { Class_Initialize(); }
						catch(Exception e)
						{
							_.SETERROR(e);
						}
					}

					public void Class_Initialize()
					{
						_.CALLm1v1(this, _env.WScript, ""Echo"", ""Here!"");
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        [TestMethod, MyFact]
        public void ClassInitializeCallHappensAfterFieldsSetToNull()
        {
            var source = @"
				CLASS C1
					PRIVATE mName
					PUBLIC SUB Class_Initialize
						mName = ""Test""
					END SUB
				END CLASS
			";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
						mName = null;
						try { Class_Initialize(); }
						catch(Exception e)
						{
							_.SETERROR(e);
						}
					}

					private object mName { get; set; }

					public void Class_Initialize()
					{
						mName = ""Test"";
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// If a GET property just returns a value then there's no need to define a return value, set it, then return that temporary reference (this is the
        /// same as for FUNCTION but not SUB, since SUB does not return a value)
        /// </summary>
        [TestMethod, MyFact]
        public void PropertyGetterThatHasSingleLineReturnsIsTranslatedIntoSingleLineReturn()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
						Name = ""C1""
					END PROPERTY
				END CLASS";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}
					[TranslatedProperty(""Name"")]
					public object Name()
					{
						return ""C1"";
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// If a GET property is NOT just a single return-value-setting-statement then a temporary return reference is declared which is set (potentially
        /// multiple times, depending upon the getter's implementation) and returned at the end. The logic to determine whether a no-temporary-reference
        /// short cut may be applied is very simplistic and only applies to the simplest cases.
        /// </summary>
        [TestMethod, MyFact]
        public void PropertyGetterThatHasMultipleLinesFollowsStandardFormat()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
						WScript.Echo ""get_Name""
						Name = ""C1""
					END PROPERTY
				END CLASS";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}
					[TranslatedProperty(""Name"")]
					public object Name()
					{
						object Name_retVal = null;
						_.CALLm1v1(this, _env.WScript, ""Echo"", ""get_Name"");
						Name_retVal = ""C1"";
						return Name_retVal;
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// LET or SET properties would seem like they should error if they try to return a value (the same as a SUB would), but for some reason VBScript just
        /// ignores the sort-of-return-value setting (it evaluates the right-hand side of the statement but doesn't return anything and doesn't error)
        /// </summary>
        [TestMethod, MyFact]
        public void NonGetPropertyIgnoresAnyReturnValueSetting()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						Name = ""C1""
					END PROPERTY
				END CLASS";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}
					[TranslatedProperty(""Name"")]
					public void Name(ref object value)
					{
						_.VAL(""C1"");
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// This is similar to NonGetPropertyIgnoresAnyReturnValueSetting, where a LET property accessor includes a value-setting statement which appears to
        /// target the current property, but that value-setting statement specifies a SET. This means that the right-hand side of the statement must be of
        /// an object reference type. This is nothing to with whether the property accessor is a LET or SET, it is solely down to whether the value-setting
        /// statement begins with "SET" or not.
        /// </summary>
        [TestMethod, MyFact]
        public void NonGetPropertyIgnoresAnyReturnValueSettingButSetSemanticsAreRespectedWhereSpecified()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						SET Name = ""C1""
					END PROPERTY
				END CLASS";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}
					[TranslatedProperty(""Name"")]
					public void Name(ref object value)
					{
						_.OBJ(""C1"");
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// This is similar to the NonGetPropertyIgnoresAnyReturnValueSetting test except that if the left-hand side of a value-setting statement within a LET
        /// or SET property specifies the name of that property WITH brackets, then it will try to call itself (potentially infinite-looping, depending upon
        /// implementation)
        /// </summary>
        [TestMethod, MyFact]
        public void NonGetPropertyCallsSelfIfBracketsAreSpecifiedAroundRecursivePropertyUpdate()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						Name() = ""C1""
					END PROPERTY
				END CLASS";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}
					[TranslatedProperty(""Name"")]
					public void Name(ref object value)
					{
						_.SETm1a0(this, this, ""Name"", ""C1"");
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// Indexed properties can not be directly represented in C# (well, one may be - the default indexed property - but it can't be explicitly named and things
        /// will fall apart if there need to be multiple indexed properties if this is the only mechanism used) so some extra logic is layered on; the properties
        /// are translated into functions and the parent class inherits TranslatedPropertyIReflectImplementation, which does some mapping work for calling code.
        /// </summary>
        [TestMethod, MyFact]
        public void IndexedPropertiesNeedSpecialLoveAndCare()
        {
            var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Blah(ByVal i, ByVal j, ByVal value)
					END PROPERTY
				END CLASS
			";
            var expected = @"
				[ComVisible(true)]
				[ClassInterface(ClassInterfaceType.AutoDispatch)]
				[SourceClassName(nameof(C1))]
				public sealed class C1 : TranslatedPropertyIReflectImplementation
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
					}

					[TranslatedProperty(""Blah"")]
					public void Blah(object i, object j, object value)
					{
					}
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

		[TestMethod]
		public void ExternalFuncX() // see 'TryToGetDeclaredReferenceDetails'
        {
            var source = @"Dim Person: set Person = GetPersonForAgent(123)";

			//NonNullImmutableList<string> testExternalDependencies = new NonNullImmutableList<string>().Add("hlmodel");
			List<ExternalMemberMethodInfo> externalMemberMethods = new List<ExternalMemberMethodInfo>();
			externalMemberMethods.Add(new ExternalMemberMethodInfo("hlmodel", "GetPersonForAgent"));
            //string actualCsCodeRaw = DefaultTranslator.TranslateWithoutScaffolding(TestCulture, source, testExternalDependencies, externalMemberMethods, []);
			var expected = @"_outer.Person = _.OBJ(_.CALLm1v1(this, hlmodel, ""GetPersonForAgent"", (Int16)123));";
			TestCSharpCodeTranslationWithoutScaffoldingX(expected, source, ["hlmodel"], externalMemberMethods, []); // SKY101
			//Assert.Inconclusive();
        }
    }
}
