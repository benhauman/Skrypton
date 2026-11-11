#nullable enable
using System.Collections.Generic;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// An in-memory IHostClipboardHostService for the tests (the user's clipboard is not touched)
    /// </summary>
    internal sealed class TestClipboardHostService : IHostClipboardHostService
    {
        public Dictionary<string, string> Formats { get; } = new Dictionary<string, string>();

        public bool SetText(string format, string text)
        {
            Formats[format] = text;
            return true;
        }

        public string? GetText(string format) => Formats.TryGetValue(format, out string? text) ? text : null;

        public bool Clear(string format)
        {
            if (format.Length == 0)
                Formats.Clear();
            else
                Formats.Remove(format);
            return true;
        }
    }
}
