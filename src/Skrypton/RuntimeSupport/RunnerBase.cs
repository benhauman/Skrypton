using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Skrypton.RuntimeSupport
{
    public abstract class RunnerBase
    {
        internal IProvideVBScriptCompatFunctionalityToIndividualRequests CompatLayer { get; } // rename it to '_'
        protected RunnerBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            CompatLayer = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }

        public abstract EnvironmentReferencesBase CreateEnvironmentReferencesInstance();

        public static RunnerBase CreateRunnerInstanceForType(Type runnerType, IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            return (RunnerBase)Activator.CreateInstance(runnerType, [compatLayer]);
        }

        public abstract GlobalReferencesBase Run(EnvironmentReferencesBase environmentReferences);
    }

    public abstract class RunnerBaseT<TEnvironmentReferences, TGlobalReferences> : RunnerBase
        where TEnvironmentReferences : EnvironmentReferencesBase, new()
        where TGlobalReferences : GlobalReferencesBase
    {
        protected RunnerBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
        }
        public override EnvironmentReferencesBase CreateEnvironmentReferencesInstance()
        {
            return new TEnvironmentReferences();
        }

        protected abstract TGlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences environmentReferences);

        public override GlobalReferencesBase Run(EnvironmentReferencesBase environmentReferences)
        {
            TGlobalReferences globalReferences = CreateGlobalReferences(CompatLayer, (TEnvironmentReferences)environmentReferences);
            Go(CompatLayer, (TEnvironmentReferences)environmentReferences, globalReferences);
            return globalReferences;
        }

        protected abstract void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences environmentReferences, TGlobalReferences globalReferences);
    }

    public abstract class EnvironmentReferencesBase
    {
        private readonly Dictionary<string, object> _externalReferences = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        protected EnvironmentReferencesBase()
        {

        }

        public void InitializeExternalReference(string referenceName, object reference)
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            _externalReferences[referenceName] = reference ?? throw new ArgumentNullException(nameof(reference)); // Use DBValue.Null for nulls.
        }

        protected object? GetExternalReferenceAsObject([CallerMemberName] string referenceName = "")
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            if (_externalReferences.TryGetValue(referenceName, out object reference))
                return reference;
            return null;//?!?
        }
        protected void RestoreExternalReferenceAsObject(object newInstance, [CallerMemberName] string referenceName = "")
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            object? current = GetExternalReferenceAsObject(referenceName);
            if (current != null && newInstance != null && current != newInstance)
            {
                throw new InvalidOperationException("not same");
            }
            else
            {
                if (newInstance == null)
                    _externalReferences.Remove(referenceName);
                else
                    _externalReferences[referenceName] = newInstance;
            }
        }
    }
    public abstract class GlobalReferencesBase
    {
        protected GlobalReferencesBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferencesBase env)
        {

        }

        //internal bool MembersCollected { get; set; }
        //private readonly Dictionary<string, bool> _methodNames = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        //internal void CollectMembers()
        //{
        //    if (!MembersCollected)
        //    {
        //        var declaringType = GetType();

        //        var mis = declaringType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly);
        //        foreach(MethodInfo mi in mis)
        //        {

        //        }


        //        MembersCollected = true;
        //    }

        //}

        internal MethodInfo GetMethodInfoByNameAndArgs(string methodName, object[] args)
        {
            var declaringType = GetType();

            var mis = declaringType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly);

            MethodInfo? candidate = null;
            foreach (MethodInfo mi in mis)
            {
                if (string.Equals(mi.Name, methodName, StringComparison.OrdinalIgnoreCase))
                {
                    if (args.Length <= mi.GetParameters().Length)
                    {
                        if (args.Length > 0)
                        {
                            // validate parameter types if their type are acceptable.
                            throw new NotSupportedException($"{mi.Name} args.Count:{args.Length}");
                        }
                        if (mi.GetParameters().Length > 0)
                        {
                            throw new NotSupportedException($"{mi.Name} prms.Count:{mi.GetParameters().Length}");
                        }



                        if (candidate == null)
                        {
                            candidate = mi;
                        }
                        else
                        {
                            if (candidate.GetParameters().Length > mi.GetParameters().Length)
                            {
                                candidate = mi; // this method accepts less arguments, take it
                            }
                        }
                    }
                }

            }

            if (candidate == null)
                throw new InvalidOperationException($"Method '{methodName}'() could not be found");
            return candidate;
        }
    }
    public abstract class GlobalReferencesBaseT<TEnvironmentReferences> : GlobalReferencesBase where TEnvironmentReferences : EnvironmentReferencesBase
    {
        protected GlobalReferencesBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences env)
            : base(compatLayer, env)
        {

        }
    }
}
