using Skrypton.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations;
using System;
using System.Diagnostics;

namespace Skrypton.Tests.Application;

internal sealed class DispatchProxyForHLHelperPFA : IReflectOnClrType
{
    public DispatchProxyForHLHelperPFA() // VBScript: CreateObject("helpline.hlcontrols.HLHelperPFA")
    {

    }

    public object GetPersonForAgent(object modelContext, int agentId)
    {
        Console.WriteLine($"[HLHelperPFA].GetPersonForAgent(agentId:{agentId})");
        if (agentId == 710)
            return new AgentPerson(agentId);
        throw new NotImplementedException($"[HLHelperPFA].GetPersonForAgent(agentId:{agentId})");
    }

    [DebuggerDisplay("PFA:{_agentId}")]
    private sealed class AgentPerson(int agentId)
    {
        private readonly int _agentId = agentId;
    }
}
