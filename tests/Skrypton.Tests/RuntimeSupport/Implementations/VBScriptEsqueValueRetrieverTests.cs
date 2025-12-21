
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Compat;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;
//using Microsoft.VisualStudio.TestTools.UnitTesting;///using Xunit//#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    // TODO: Add a test that complements ExpressionGeneratorTests.PropertyAccessOnStringLiteralResultsInRuntimeError, that confirms that if a string value is
    // provided as call target that an error will be raised for property or method access attempts
    [TestClass]
    public class VBScriptEsqueValueRetrieverTests : TestBase
    {
        [TestMethod, MyFact]
        public void VALDoesNotAlterNull()
        {
            myAssert.AreEqual(
                null,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(null)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterOne()
        {
            myAssert.AreEqual(
                1,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(1)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterOnePointOneFloat()
        {
            myAssert.AreEqual(
                1.1f,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(1.1f)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterOnePointOneDouble()
        {
            myAssert.AreEqual(
                1.1d,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(1.1d)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterOnePointOneDecimal()
        {
            myAssert.AreEqual(
                1.1m,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(1.1m)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterMinusOne()
        {
            myAssert.AreEqual(
                -1,
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(-1)
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterEmptyString()
        {
            myAssert.AreEqual(
                "",
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL("")
            );
        }

        [TestMethod, MyFact]
        public void VALDoesNotAlterNonEmptyString()
        {
            myAssert.AreEqual(
                "Test",
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL("Test")
            );
        }

        [TestMethod, MyFact]
        public void VALFailsOnTranslatedClassWithNoDefaultMember()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new Translatedclasswithnodefaultmember()));
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new Translatedclasswithnodefaultmember()));
        }

        [SourceClassName("TranslatedClassWithNoDefaultMember")]
        private class Translatedclasswithnodefaultmember { }

        [TestMethod, MyFact]
        public void VALFailsOnComObjectWithNoParameterlessDefaultMember()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            var dictionary = Activator.CreateInstance(typeof(MyScriptingDictionary));//lubo: Type.GetTypeFromProgID("Scripting.Dictionary"));
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(dictionary));
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(dictionary));
        }

        /// <summary>
        /// This test relates to a fix just applied to GenerateSetInvoker (where an argumentsArray reference was being used instead of invokeArguments, which meant that the same set of arguments
        /// were being reused on each call)
        /// </summary>
        [TestMethod, MyFact]
        public void EnsureThatOldArgumentsAreNotReusedInSubsequentIDispatchCalls()
        {
            // This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
            var dict = Activator.CreateInstance(typeof(MyScriptingDictionary));//lubo: Type.GetTypeFromProgID("Scripting.Dictionary"));
            using (var _ = Skrypton.RuntimeSupport.DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
            {
                _.SET(1, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("a"));
                _.SET(2, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("b"));
                myAssert.AreEqual(2, _.CALL(context: null, target: dict, member1: "Count"));
            }
        }

        [TestMethod, MyFact]
        public void VALFailsOnNonComVisibleNonTranslatedClasses()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new NonComVisibleNonTranslatedClass()));
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new NonComVisibleNonTranslatedClass()));
        }

        private class NonComVisibleNonTranslatedClass { }

        [TestMethod, MyFact]
        public void VALSupportsIsDefaultAttributeOnTranslatedClasses()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            myAssert.AreEqual("name!", DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new translatedclasswithdefaultmember()));
            myAssert.AreEqual("name!", DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new translatedclasswithdefaultmember()));
        }

        [SourceClassName("TranslatedClassWithNoDefaultMember")]
        private class translatedclasswithdefaultmember
        {
            [IsDefault]
            public string name() { return "name!"; }
        }

        [TestMethod, MyFact]
        public void VALSupportsDefaultMemberAttributeOnComVisibleNonTranslatedClasses()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            myAssert.AreEqual("name!", DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithDefaultMember()));
            myAssert.AreEqual("name!", DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithDefaultMember()));
        }

        [ComVisible(true)]
        [DefaultMember("Name")]
        private class ComVisibleNonTranslatedClassWithDefaultMember
        {
            public string Name { get { return "name!"; } }
        }

        [TestMethod, MyFact]
        public void VALSupportsToStringOnComVisibleNonTranslatedClasses()
        {
            // Execute twice to ensure that the TryVAL caching does not affect the result
            var target = new ComVisibleNonTranslatedClassWithDefaultMember();
            myAssert.AreEqualString(
                "Skrypton.Tests.RuntimeSupport.Implementations.VBScriptEsqueValueRetrieverTests+ComVisibleNonTranslatedClassWithNoDefaultMember",
                (string)DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithNoDefaultMember())
            );
            myAssert.AreEqualString(
                "Skrypton.Tests.RuntimeSupport.Implementations.VBScriptEsqueValueRetrieverTests+ComVisibleNonTranslatedClassWithNoDefaultMember",
                (string)DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithNoDefaultMember())
            );
        }

        [ComVisible(true)]
        private class ComVisibleNonTranslatedClassWithNoDefaultMember { }

        [TestMethod, MyFact]
        public void IFOfNullIsFalse()
        {
            myAssert.False(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(null)
            );
        }

        [TestMethod, MyFact]
        public void IFOfZeroIsFalse()
        {
            myAssert.False(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(0)
            );
        }

        [TestMethod, MyFact]
        public void IFOfOneIsTrue()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(1)
            );
        }

        [TestMethod, MyFact]
        public void IFOfMinusOneIsTrue()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(-1)
            );
        }

        [TestMethod, MyFact]
        public void IFOfOnePointOneIsTrue()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(1.1)
            );
        }

        /// <summary>
        /// VBScript doesn't round the number down to zero and find 0.1 to be false, it just checks that the number is non-zero
        /// </summary>
        [TestMethod, MyFact]
        public void IFOfPointOneIsTrue()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF(0.1)
            );
        }

        [TestMethod, MyFact]
        public void IFOfPointStringRepresentationOfOneIsTrue()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF("1")
            );
        }

        [TestMethod, MyFact]
        public void IFThrowsExceptionForBlanksString()
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF("");
            });
        }

        [TestMethod, MyFact]
        public void IFThrowsExceptionForNonNumericString()
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF("one");
            });
        }

        [TestMethod, MyFact]
        public void IFIgnoresWhiteSpaceWhenParsingStrings()
        {
            myAssert.True(
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.IF("   1    ")
            );
        }

        [TestMethod, MyFact]
        public void NativeClassSupportsMethodCallWithArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(
                new PseudoField { value = "value:F1" },
                _.CALL(
                    context: null,
                    target: new PseudoRecordset(),
                    members: new[] { "fields" },
                    argumentProviderBuilder: _.ARGS.Val("F1")
                ),
                new PseudoFieldObjectComparer()
            );
        }

        [TestMethod, MyFact]
        public void NativeClassSupportsDefaultMethodCallWithArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(
                new PseudoField { value = "value:F1" },
                _.CALL(
                    context: null,
                    target: new PseudoRecordset(),
                    members: new string[0],
                    argumentProviderBuilder: _.ARGS.Val("F1")
                ),
                new PseudoFieldObjectComparer()
            );
        }
        /*
                //lubo[TestMethod, MyFact]
                public void ADORecordsetSupportsNamedFieldAccess()
                {
                    var recordset = new ADODB.Recordset();
                    recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
                    recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
                    recordset.AddNew();
                    recordset.Fields["name"].Value = "TestName";
                    recordset.Update();

                    var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
                    myAssert.AreEqual(
                        recordset.Fields["name"],
                        _.CALL(
                            context: null,
                            target: recordset,
                            members: new[] { "fields" },
                            argumentProviderBuilder: _.ARGS.Val("name")
                        ),
                        new ADOFieldObjectComparer()
                    );
                }

                //lubo [TestMethod, MyFact]
                public void ADORecordsetSupportsDefaultFieldAccess()
                {
                    var recordset = new ADODB.Recordset();
                    recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
                    recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
                    recordset.AddNew();
                    recordset.Fields["name"].Value = "TestName";
                    recordset.Update();

                    var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
                    myAssert.AreEqual(
                        recordset.Fields["name"],
                        _.CALL(
                            context: null,
                            target: recordset,
                            members: new string[0],
                            argumentProviderBuilder: _.ARGS.Val("name")
                        ),
                        new ADOFieldObjectComparer()
                    );
                }

                /// <summary>
                /// This describes an extremely common VBScript pattern - rstResults("name") needs to return a value by using the default Fields access
                /// (passing through "name" to it) and then default access of the Field's Value property (which requires a VAL call, which must be
                /// included in translated code if a value type is expected - any time other than when a SET statement is present as part of a
                /// variable assignment)
                /// </summary>
                // lubo[TestMethod, MyFact]
                public void ADORecordsetSupportsDefaultFieldValueAccess()
                {
                    var recordset = new ADODB.Recordset();
                    recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
                    recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
                    recordset.AddNew();
                    recordset.Fields["name"].Value = "TestName";
                    recordset.Update();

                    var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
                    myAssert.AreEqual(
                        "TestName",
                        _.VAL(
                            _.CALL(
                                context: null,
                                target: recordset,
                                members: new string[0],
                                argumentProviderBuilder: _.ARGS.Val("name")
                            )
                        )
                    );
                }
        */
        [TestMethod, MyFact]
        public void OneDimensionalArrayAccessIsSupported()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            var data = new object[] { "One" };
            myAssert.AreEqual(
                "One",
                _.CALL(
                    context: null,
                    target: data,
                    members: new string[0],
                    argumentProviderBuilder: _.ARGS.Val("0")
                )
            );
        }

        [TestMethod, MyFact]
        public void ByRefArgumentIsUpdatedAfterCall()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            object arg0 = 1;
            _.CALL(context: null, target: this, member1: "ByRefArgUpdatingFunction", argumentProviderBuilder: _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(false));
            myAssert.AreEqual(123, arg0);
        }

        [TestMethod, MyFact]
        public void ByRefArgumentIsUpdatedAfterCallEvenIfExceptionIsThrown()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            object arg0 = 1;
            try
            {
                _.CALL(context: null, target: this, member1: "ByRefArgUpdatingFunction", argumentProviderBuilder: _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(true));
            }
            catch { }
            myAssert.AreEqual(123, arg0);
        }

        [TestMethod, MyFact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithZeroValues()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(0, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS));
        }

        [TestMethod, MyFact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithSingleValue()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(1, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS.Val(1)));
        }

        [TestMethod, MyFact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithTwoValues()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(2, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS.Val(1).Val(2)));
        }

        public void ByRefArgUpdatingFunction(ref object arg0, bool throwExceptionAfterUpdatingArgument)
        {
            arg0 = 123;
            if (throwExceptionAfterUpdatingArgument)
                throw new Exception("Example exception");
        }

        public Nullable<int> GetNumberOfArgumentsPassedInParamsObjectArray(params object[] args)
        {
            return (args == null) ? (int?)null : args.Length;
        }

        /// <summary>
        /// When a CALL execution is generated by the translator, the string member accessors should not be renamed - to avoid C# keywords, for example.
        /// If the target is a translated class then any members that use reserved C# keywords must be renamed, but if the target is not a translated
        /// class (a COM component, for example) then the call will fail, so the renaming must not be done at translation time. This means that the
        /// CALL implementation must support name rewriting for when the target is a translated-from-VBScript C# class.
        /// </summary>
        [TestMethod, MyFact]
        public void StringMemberAccessorValuesShouldNotBeRewrittenAtTranslationTime()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(
                "Success!",
                _.CALL(context: null, target: new ImpressionOfTranslatedClassWithRewrittenPropertyName(), member1: "Params")
            );
        }

        [TestMethod, MyFact]
        public void DelegateWithIncorrectNumberOfArguments()
        {
            var parameterLessDelegate = (Func<object>)(() => "delegate result");
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<TargetParameterCountException>(
                () => _.CALL(context: null, target: parameterLessDelegate, members: new string[0], argumentProvider: _.ARGS.Val(1).GetArgs())
            );
        }

        /// <summary>
        /// In C#, it's fine to access an index within a string since a string is an array of characters. But in VBScript, it's not.
        /// </summary>
        [TestMethod, MyFact]
        public void ItIsNotValidToAccessStringValueWithArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<TypeMismatchException>(
                () => _.CALL(context: null, target: "abc", members: new string[0], argumentProvider: _.ARGS.Val(0).GetArgs())
            );
        }

        /// <summary>
        /// DispId(0) is only supported when the match is unambiguous - previously, DispId(0) being specified on a property and on the getter for
        /// that property was considered an ambiguous match, but that shouldn't be the case since they both effectively refer to the same thing
        /// </summary>
        [TestMethod, MyFact]
        public void SupportDispIdoZeroBeingRepeatedOnPropertyAndOnPropertyGetterWhenDefaultMemberRequired()
        {
            var target = new DispIdZeroRepeatedOnPropertyAndItsGetter("test");
            myAssert.AreEqual(
                "test",
                DefaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever.VAL(target)
            );
        }

        [TestMethod, MyFact]
        public void DispIdZeroPropertySettingWorksWithValueTypes()
        {
            // This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
            var dict = Activator.CreateInstance(typeof(MyScriptingDictionary));//lubo: Type.GetTypeFromProgID("Scripting.Dictionary"));
            var valueTypeValueToRecord = 123;
            using (var _ = DefaultRuntimeSupportClassFactoryInstance.Get())
            {
                _.SET(valueTypeValueToRecord, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("ACCO"));
            }
        }

        [TestMethod, MyFact]
        public void DispIdZeroPropertySettingWorksWithReferenceTypes()
        {
            // This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
            var dict = Activator.CreateInstance(typeof(MyScriptingDictionary));//lubo: Type.GetTypeFromProgID("Scripting.Dictionary"));
            var referenceTypeValueToRecord = Activator.CreateInstance(typeof(MyScriptingDictionary));//lubo: Type.GetTypeFromProgID("Scripting.Dictionary"));
            using (var _ = DefaultRuntimeSupportClassFactoryInstance.Get())
            {
                _.SET(referenceTypeValueToRecord, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("ACCO"));
            }
        }

        [ComVisible(true)]
        private class DispIdZeroRepeatedOnPropertyAndItsGetter
        {
            private readonly string _name;
            public DispIdZeroRepeatedOnPropertyAndItsGetter(string name)
            {
                _name = name;
            }

            [DispId(0)]
            public string Name
            {
                [DispId(0)]
                get { return _name; }
            }
        }


        [TestMethod, MyFact]
        public void CallPrivateMemberFromWithinContextOfClassShouldWork()
        {
            const string name = "test";
            var classWithPrivateMember = new ClassWithPrivateGetNameMethod(name);
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(
                name,
                _.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "GetName")
            );
        }

        [TestMethod, MyFact]
        public void CallPrivateMemberFromOutsideContextOfClassShouldThrow()
        {
            const string name = "test";
            var classWithPrivateMember = new ClassWithPrivateGetNameMethod(name);
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<MissingMemberException>(() =>
                _.CALL(context: null, target: classWithPrivateMember, member1: "GetName")
            );
        }

        [TestMethod, MyFact]
        public void SettingPrivatePropertyFromWithinContextOfClassShouldWork()
        {
            const string name = "test";
            var classWithPrivateMember = new ClassWithPrivateNameProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Name");
            myAssert.AreEqual(
                name,
                _.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Name")
            );
        }

        [TestMethod, MyFact]
        public void SettingPrivatePropertyFromOutsideContextOfClassShouldThrow()
        {
            const string name = "test";
            var classWithPrivateMember = new ClassWithPrivateNameProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<MissingMemberException>(() =>
                _.SET(name, context: null, target: classWithPrivateMember, optionalMemberAccessor: "Name")
            );
        }

        private class ClassWithPrivateGetNameMethod
        {
            private readonly string _name;
            public ClassWithPrivateGetNameMethod(string name)
            {
                _name = name;
            }

            private string GetName()
            {
                return _name;
            }
        }

        private class ClassWithPrivateNameProperty // This is approximately how a VBScript property is translated (into getter and setter methods)
        {
            private object mname { get; set; }
            private object name()
            {
                return mname;
            }
            private void name(object strname)
            {
                mname = strname;
            }
        }

        [TestMethod, MyFact]
        public void SettingPrivateIndexedPropertyFromWithinContextOfClassShouldWork()
        {
            const string name = "test";
            var i = new object();
            var classWithPrivateMember = new ClassWithPrivateIndexedProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
            myAssert.AreEqual(
                name,
                _.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
            );
        }

        [TestMethod, MyFact]
        public void SettingPrivateIndexedPropertyFromOutsideContextOfClassShouldThrow()
        {
            const string name = "test";
            var i = new object();
            var classWithPrivateMember = new ClassWithPrivateIndexedProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<MissingMethodException>(() =>
                _.SET(name, context: null, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i))
            );
        }

        // This is approximately how a VBScript indexed property is translated (into getter and setter methods) since C# only supports a single indexed property
        private class ClassWithPrivateIndexedProperty : TranslatedPropertyIReflectImplementation
        {
            private Dictionary<object, object> _values = new Dictionary<object, object>();

            [TranslatedProperty("Test")]
            private object test(object i)
            {
                return _values.ContainsKey(i) ? _values[i] : null;
            }

            [TranslatedProperty("Test")]
            private void test(object i, object value)
            {
                _values[i] = value;
            }
        }

        [TestMethod, MyFact]
        public void SettingPublicIndexedPropertyFromWithinContextOfClassShouldWork()
        {
            const string name = "test";
            var i = new object();
            var classWithPrivateMember = new ClassWithPublicIndexedProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
            myAssert.AreEqual(
                name,
                _.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
            );
        }

        [TestMethod, MyFact]
        public void SettingPublicIndexedPropertyFromOutsideContextOfClassShouldWork()
        {
            const string name = "test";
            var i = new object();
            var classWithPrivateMember = new ClassWithPublicIndexedProperty();
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
            myAssert.AreEqual(
                name,
                _.CALL(context: null, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
            );
        }

        // This is approximately how a VBScript indexed property is translated (into getter and setter methods) since C# only supports a single indexed property
        private class ClassWithPublicIndexedProperty : TranslatedPropertyIReflectImplementation
        {
            private Dictionary<object, object> _values = new Dictionary<object, object>();

            [TranslatedProperty("Test")]
            public object test(object i)
            {
                return _values.ContainsKey(i) ? _values[i] : null;
            }

            [TranslatedProperty("Test")]
            public void test(object i, object value)
            {
                _values[i] = value;
            }
        }

        [TestMethod, MyFact]
        public void ByRefIndexArgumentOnPublicPropertySetterShouldAcceptUpdatesWhenCalledOverIReflect()
        {
            object i = "123";
            object value = "xyz";
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(
                value,
                context: null,
                target: new ClassWithPublicIndexedPropertyThatHasByRefArguments(),
                optionalMemberAccessor: "Test",
                argumentProviderBuilder: _.ARGS.Ref(i, iUpdate => { i = iUpdate; })
            );
            myAssert.AreEqual(i, "456");
        }

        private class ClassWithPublicIndexedPropertyThatHasByRefArguments : TranslatedPropertyIReflectImplementation
        {
            [TranslatedProperty("Test")]
            public void test(ref object i, ref object value)
            {
                i = "456";
            }
        }

        [TestMethod, MyFact]
        public void ByRefIndexArgumentOnPublicPropertySetterShouldAcceptUpdatesWhenNotCalledOverIReflect()
        {
            object i = "123";
            object value = "xyz";
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            _.SET(
                value,
                context: null,
                target: new ClassWithPublicIndexedPropertyThatHasByRefArgumentsButThatIsNotCalledOverIReflect(),
                optionalMemberAccessor: "Test",
                argumentProviderBuilder: _.ARGS.Ref(i, iUpdate => { i = iUpdate; })
            );
            myAssert.AreEqual(i, "456");
        }

        // Classes translated from VBScript that have indexed properties will be derived from TranslatedPropertyIReflectImplementation but we need to test the
        // logic when translated-from-VBScript code calls into not-translated-from-VBScript code as well (to ensure that the indexed arguments are ByRef-updated)
        private class ClassWithPublicIndexedPropertyThatHasByRefArgumentsButThatIsNotCalledOverIReflect
        {
            public void test(ref object i, ref object value)
            {
                i = "456";
            }
        }
        /*
                //lubo[TestMethod, MyFact]
                public void CallingCLRMethodsThatHaveValueTypeParametersWorksWithReferenceTypes()
                {
                    var recordset = new ADODB.Recordset();
                    recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
                    recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
                    recordset.AddNew();
                    recordset.Fields["name"].Value = "TestName";
                    recordset.Update();

                    var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;

                    object objField = _.CALL(
                        context: null,
                        target: recordset,
                        members: new string[0],
                        argumentProviderBuilder: _.ARGS.Val("name")
                    );

                    myAssert.AreEqual(
                        "TestName",
                        _.CALL(
                            context: null,
                            target: this,
                            member1: "MockMethodReturningInputString",
                            argumentProviderBuilder: _.ARGS.Ref(objField, v => { objField = v; })
                        )
                    );
                }
        */
        public string MockMethodReturningInputString(string input)
        {
            return input;
        }

        [TestMethod, MyFact]
        public void NothingShouldBeReturnedForNullForPropertyOfComVisibleType()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            var value = _.CALL(context: null, target: new ClassWithComVisiblePropertyThatIsAlwaysNull(), member1: "Value");
#pragma warning disable CA1416 // Validate platform compatibility
            myAssert.IsType<DispatchWrapper>(value);
            myAssert.Null(((DispatchWrapper)value).WrappedObject);
#pragma warning restore CA1416 // Validate platform compatibility
        }

        [TestMethod, MyFact]
        public void NothingShouldNotBeReturnedForNullForPropertyOfObjectType()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Null(
                _.CALL(context: null, target: new ClassWithObjectPropertyThatIsAlwaysNull(), member1: "Value")
            );
        }

        /// <summary>
        /// VBScript has its own ideas about what constitutes a value type, so it won't get Nothing from a null property value if the property's type is string (even though
        /// string is ComVisible and not "object" and not a .NET value type)
        /// </summary>
        [TestMethod, MyFact]
        public void NothingShouldNotBeReturnedForNullForPropertyOfStringType()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Null(
                _.CALL(context: null, target: new ClassWithObjectPropertyThatIsAlwaysNull(), member1: "Value")
            );
        }

        [ComVisible(true)]
        private class ClassWithComVisiblePropertyThatIsAlwaysNull
        {
            public ClassWithComVisiblePropertyThatIsAlwaysNull Value { get { return null; } }
        }

        [ComVisible(true)]
        private class ClassWithObjectPropertyThatIsAlwaysNull
        {
            public object Value { get { return null; } }
        }

        /// <summary>
        /// When a VBScript WSC has a reference to a ComVisible object with a DispId zero method that has a single [Optional] parameter and it wants to coerce that object into
        /// a value type, it will call the default member and pass a Missing value to the argument. The VBScriptTranslator runtime library has not previously done this - this
        /// test illustrates the issue.
        /// </summary>
        [TestMethod, MyFact]
        public void WhenLookingForParameterLessDefaultMemberOnComVisibleClassSupportOptionalArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(
                "YEAH!",
                _.VAL(_.CALL(context: null, target: new ClassWithDefaultMethodWithSingleOptionalArgument()))
            );
        }

        [ComVisible(true)]
        public sealed class ClassWithDefaultMethodWithSingleOptionalArgument
        {
            [DispId(0)]
            public object Item([Optional] object arg)
            {
                return "YEAH!";
            }
        }

        [TestMethod, MyTheory, MyMemberData(nameof(ZeroArgumentBracketSuccessData))]
        public void ZeroArgumentBracketSuccessCases(string description, object target, string[] memberAccessors, bool useBracketsWhereZeroArguments, object expectedResult)
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            var args = _.ARGS;
            if (useBracketsWhereZeroArguments)
                args = args.ForceBrackets();
            myAssert.AreEqual(expectedResult, _.CALL(context: null, target: target, members: memberAccessors, argumentProvider: args.GetArgs()));
        }

        [TestMethod, MyTheory, MyMemberData("ZeroArgumentBracketFailData")]
        public void ZeroArgumentBracketFailCases(string description, object target, string[] memberAccessors, bool useBracketsWhereZeroArguments, Type exceptionType)
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            var args = _.ARGS;
            if (useBracketsWhereZeroArguments)
                args = args.ForceBrackets();
            myAssert.ThrowsX(exceptionType, () => _.CALL(context: null, target: target, members: memberAccessors, argumentProvider: args.GetArgs()));
        }

        public static IEnumerable<object[]> ZeroArgumentBracketSuccessData
        {
            get
            {
                var array = new object[] { 123 };
                yield return new object[] { "Array with no member accessors, properties or zero-argument brackets", array, new string[0], false, array };
                yield return new object[] { "String with no member accessors, properties or zero-argument brackets", "123", new string[0], false, "123" };

                var parameterLessDelegate = (Func<object>)(() => "delegate result");
                yield return new object[] { "Delegate with no member accessors, properties or zero-argument brackets", parameterLessDelegate, new string[0], false, parameterLessDelegate };
                yield return new object[] { "Delegate with no member accessors, properties WITH zero-argument brackets", parameterLessDelegate, new string[0], true, "delegate result" };

                yield return new object[] { "VBScript class property without brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "Name" }, false, "test" };
                yield return new object[] { "VBScript class property WITH brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "Name" }, true, "test" };
                yield return new object[] { "VBScript class function without brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "GetName" }, false, "test" };
                yield return new object[] { "VBScript class function WITH brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "GetName" }, true, "test" };

                yield return new object[] { "COM component property without brackets", Activator.CreateInstance(typeof(MyScriptingDictionary)), new[] { "Count" }, false, 0 };
            }
        }

        public static IEnumerable<object[]> ZeroArgumentBracketFailData
        {
            get
            {
                yield return new object[] { "String with zero-argument brackets", "123", new string[0], true, typeof(TypeMismatchException) };
                yield return new object[] { "Array with zero-argument brackets", new object[] { 123 }, new string[0], true, typeof(SubscriptOutOfRangeException) };
                yield return new object[] { "COM component property with brackets", Activator.CreateInstance(typeof(MyScriptingDictionary)), new[] { "Count" }, true, typeof(IDispatchAccess.IDispatchAccessException) };
                yield return new object[] { "Delegate with a member accessors", (Func<object>)(() => "delegate result"), new[] { "Name" }, false, typeof(ArgumentException) };
            }
        }

        /// <summary>
        /// This will be used in tests that target "VBScript classes", meaning classes translated from VBScript into C# (as opposed to, say, COM components)
        /// </summary>
        private class ZeroArgumentBracketExampleClass
        {
            public ZeroArgumentBracketExampleClass(string name) { Name = name; }
            public string Name { get; private set; }
            public string GetName() { return Name; }
        }

        private class ImpressionOfTranslatedClassWithRewrittenPropertyName
        {
            /// <summary>
            /// This is a property that would have been rewritten from one named "Param" in VBScript (which is valid in VBScript but is a reserved
            /// keyword in C# and so may not appear in C# without some manipulation)
            /// </summary>
            public object rewritten_params { get { return "Success!"; } }
        }

        [TestMethod, MyTheory, MyMemberData("AcceptableEnumerableValueData")]
        public void AcceptableEnumerableValueCases(string description, object value, IEnumerable<object> expectedResults)
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.AreEqual(expectedResults, _.ENUMERABLE(value).Cast<Object>()); // Cast to Object because we care about testing the contents, not the element type
        }

        [TestMethod, MyTheory, MyMemberData("UnacceptableEnumerableValueData")]
        public void UnacceptableEnumerableValueCases(string description, object value)
        {
            var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).DefaultVBScriptValueRetriever;
            myAssert.Throws<ObjectNotCollectionException>(() => _.ENUMERABLE(value));
        }

        public static IEnumerable<object[]> AcceptableEnumerableValueData
        {
            get
            {
                yield return new object[] { "An object array", new object[] { 1, 2 }, new object[] { 1, 2 } };

                dynamic dictionary = Activator.CreateInstance(typeof(MyScriptingDictionary));
                dictionary.Add("key1", "value1");
                dictionary.Add("key2", "value2");
                yield return new object[] { "Scripting Dictionary COM component", dictionary, new object[] { "key1", "key2" } };

                var customIntEnumerable = new CustomEnumerable<int>(new[] { 1, 2, 3 });
                yield return new object[] { "Object with int values that has a valid GetEnumerator but does not implement IEnumerable", customIntEnumerable, new object[] { 1, 2, 3 } };

                var customIntEnumerableWithStructEnumerator = new CustomEnumerableWithStructEnumerator<int>(new[] { 1, 2, 3 });
                yield return new object[] { "Object with int values that has a valid GetEnumerator (that returns a struct) but does not implement IEnumerable", customIntEnumerableWithStructEnumerator, new object[] { 1, 2, 3 } };

                var customStringEnumerable = new CustomEnumerable<string>(new[] { "a", "b", "c" });
                yield return new object[] { "Object with string values that has a valid GetEnumerator but does not implement IEnumerable", customStringEnumerable, new object[] { "a", "b", "c" } };
            }
        }

        public static IEnumerable<object[]> UnacceptableEnumerableValueData
        {
            get
            {
                yield return new object[] { "Empty", null };
                yield return new object[] { "Null", DBNull.Value };
                yield return new object[] { "A string", "abc" }; // String ARE enumerable in C# but must not be treated so when mimicking VBScript
            }
        }

        private class ADOFieldObjectComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y)
            {
                if (x == null)
                    throw new ArgumentNullException("x");
                if (y == null)
                    throw new ArgumentNullException("y");
                //var fieldX = x as ADODB.Field;
                //if (fieldX == null)
                //	throw new ArgumentException("x is not an ADODB.Field");
                //var fieldY = y as ADODB.Field;
                //if (fieldY == null)
                //	throw new ArgumentException("y is not an ADODB.Field");
                //return fieldX.Value == fieldY.Value;
                throw new NotImplementedException();// lubo
            }

            public int GetHashCode(object obj)
            {
                return 0;
            }
        }

        private class PseudoFieldObjectComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y)
            {
                if (x == null)
                    throw new ArgumentNullException("x");
                if (y == null)
                    throw new ArgumentNullException("y");
                var fieldX = x as PseudoField;
                if (fieldX == null)
                    throw new ArgumentException("x is not a PseudoField");
                var fieldY = y as PseudoField;
                if (fieldY == null)
                    throw new ArgumentException("y is not a PseudoField");
                return (fieldX.value as string) == (fieldY.value as string);
            }

            public int GetHashCode(object obj)
            {
                return 0;
            }
        }

        private class PseudoRecordset
        {
            [IsDefault]
            public object fields(string fieldName)
            {
                return new PseudoField { value = "value:" + fieldName };
            }
        }

        private class PseudoField
        {
            [IsDefault]
            public object value { get; set; }
        }

        private sealed class CustomEnumerable<T>
        {
            private readonly IEnumerable<T> _values;
            public CustomEnumerable(IEnumerable<T> values)
            {
                if (values == null)
                    throw new ArgumentNullException(nameof(values));
                _values = values;
            }
            public MyEnumerator<T> GetEnumerator() { return new MyEnumerator<T>(_values); }
        }

        private sealed class CustomEnumerableWithStructEnumerator<T>
        {
            private readonly IEnumerable<T> _values;
            public CustomEnumerableWithStructEnumerator(IEnumerable<T> values)
            {
                if (values == null)
                    throw new ArgumentNullException(nameof(values));
                _values = values;
            }
            public MyStructEnumerator<T> GetEnumerator() { return new MyStructEnumerator<T>(new MyEnumerator<T>(_values)); }
        }

        private struct MyStructEnumerator<T>
        {
            private readonly MyEnumerator<T> _enumerator;
            public MyStructEnumerator(MyEnumerator<T> enumerator)
            {
                _enumerator = enumerator;
            }
            public T Current { get { return _enumerator.Current; } }
            public bool MoveNext() { return _enumerator.MoveNext(); }
            public void Reset() { _enumerator.Reset(); }
        }

        private sealed class MyEnumerator<T>
        {
            private readonly T[] _values;
            private int _index;
            public MyEnumerator(IEnumerable<T> values)
            {
                if (values == null)
                    throw new ArgumentNullException(nameof(values));
                _values = values.ToArray();
                _index = -1;
            }
            public T Current
            {
                get
                {
                    if (_index == -1)
                        throw new InvalidOperationException("Enumeration has not started");
                    return _values[_index];
                }
            }
            public bool MoveNext()
            {
                if (_index == _values.Length - 1)
                    return false;

                _index++;
                return true;
            }
            public void Reset()
            {
                _index = -1;
            }
        }
    }
}
