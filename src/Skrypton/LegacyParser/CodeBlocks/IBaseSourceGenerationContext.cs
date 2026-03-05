using System;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;

namespace Skrypton.LegacyParser.CodeBlocks;

public interface IBaseSourceGenerationContext
{
    SourceRendering.ISourceIndentHandler Indenter { get; }

    IBaseSourceGenerationContext NullIndenter();
    IBaseSourceGenerationContext Increase();
    IBaseSourceGenerationContext Decrease();
    string Indent { get; }
}

public sealed class BaseSourceGenerationContextDefault : IBaseSourceGenerationContext
{
    public static IBaseSourceGenerationContext CreateBaseSourceGenerationContext() => new BaseSourceGenerationContextDefault(new Skrypton.LegacyParser.CodeBlocks.SourceRendering.SourceIndentHandler());
    private BaseSourceGenerationContextDefault(SourceRendering.ISourceIndentHandler indenter)
    {
        Indenter = indenter ?? throw new ArgumentNullException(nameof(indenter));
    }
    public SourceRendering.ISourceIndentHandler Indenter { get; }

    public IBaseSourceGenerationContext NullIndenter() => new BaseSourceGenerationContextDefault(SourceRendering.NullIndenter.Instance);
    public IBaseSourceGenerationContext Increase() => new BaseSourceGenerationContextDefault(Indenter.Increase());
    public IBaseSourceGenerationContext Decrease() => new BaseSourceGenerationContextDefault(Indenter.Decrease());
    public string Indent => Indenter.Indent;
}