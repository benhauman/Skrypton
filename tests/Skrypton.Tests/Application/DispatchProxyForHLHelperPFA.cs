using Skrypton.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations;
using System;
using System.Diagnostics;
using System.Globalization;

namespace Skrypton.Tests.Application;

internal sealed class DispatchProxyForHLHelperPFA : IReflectOnClrType
{
    private readonly CultureInfo _culture;

    public DispatchProxyForHLHelperPFA(CultureInfo culture) // VBScript: CreateObject("helpline.hlcontrols.HLHelperPFA")
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
    }

    public object GetPersonForAgent(object modelContext, int agentId)
    {
        Console.WriteLine($"[HLHelperPFA].GetPersonForAgent(agentId:{agentId})");
        if (agentId == 710)
        {
            return new HLObjectInstance().InitializeObjectInstance(isNew: false, _culture)
                .RegisterValueKey<string>("PersonGeneralTrumpf.Responsibility", 0, 0, "ResponsibilityBSZDitzingen")
                ;
            //return new AgentPerson(agentId);
        }
        throw new NotImplementedException($"[HLHelperPFA].GetPersonForAgent(agentId:{agentId})");
    }

    //[DebuggerDisplay("PFA:{_agentId}")]
    //private sealed class AgentPerson(int agentId)
    //{
    //    private readonly int _agentId = agentId;
    //}
}
