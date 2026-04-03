using System;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Application;

internal sealed class TestScriptControlConfiguration : ScriptControlConfiguration
{
    private readonly TestBaseX _tst;

    public TestScriptControlConfiguration(TestBaseX tst, bool tempEnabled, string[] translationSuppression, string[] noWarn) : base(tempEnabled, tst.TestRunResultsDirectory, enabledLoadFromDisk: false, translationSuppression: translationSuppression, noWarn: noWarn)
    {
        _tst = tst ?? throw new ArgumentNullException(nameof(tst));
    }

    protected override void OnTempFileAdd(string filePath)
    {
        _tst.AddResultFile(filePath);
    }
}