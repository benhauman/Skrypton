using System;

namespace Skrypton.CSharpWriter.Logging
{
    public sealed class ConsoleLogger : ILogInformation
    {
        public void Warning(string content)
        {
            Console.WriteLine(content);
        }
    }
}
