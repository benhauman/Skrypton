#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

// The tests run the command line tool (and cscript) as processes, which are independent - run them in parallel (issue #134)
[assembly: Parallelize(Workers = 0, Scope = ExecutionScope.MethodLevel)]

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// The state of the machine that tests share (they run in parallel): the clipboard
    /// </summary>
    internal static class SharedMachineState
    {
        public static readonly object Clipboard = new object();
    }
}
