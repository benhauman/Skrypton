using Skrypton.Tests.RuntimeSupport.Implementations;
using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    public sealed class DialogGuiLabelControl : DialogGuiControlBase
    {
        public string Text { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
        //private ControlFont _font = ControlFont.Default;
        //public ControlFont Font
        //{
        //    get => _font;
        //    set => _font = value ?? ControlFont.Default;
        //}
    }

    [ComVisible(true)]
    internal sealed class ControlFont : IReflectOnClrType
    {
        //public ControlFont()
        //{
        //}
        //public static ControlFont Default = new ControlFont();
    }
}