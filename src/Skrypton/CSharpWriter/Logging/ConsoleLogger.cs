using System;

namespace Skrypton.CSharpWriter.Logging
{
    public class ConsoleLogger : ILogInformation
    {
        public void Warning(string content)
        {
            Console.WriteLine(content);
        }
    }
}
