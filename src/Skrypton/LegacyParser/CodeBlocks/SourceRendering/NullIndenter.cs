namespace Skrypton.LegacyParser.CodeBlocks.SourceRendering
{
    internal sealed class NullIndenter : ISourceIndentHandler
    {
        private static readonly NullIndenter _instance = new NullIndenter();
        public static NullIndenter Instance { get { return _instance; } }
        private NullIndenter() { }

        public ISourceIndentHandler Increase()
        {
            return NullIndenter.Instance;
        }

        public ISourceIndentHandler Decrease()
        {
            return NullIndenter.Instance;
        }

        public string Indent => string.Empty;
    }
}
