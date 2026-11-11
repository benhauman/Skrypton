
namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in htmlfile (parentWindow.clipboardData) reads and writes the clipboard through. Whose clipboard a script
    /// may use (the user's, a per-session one, none) is a decision for the host, so there is no default: without a registration the clipboard
    /// members raise an error.
    /// </summary>
    public interface IHostClipboardHostService
    {
        /// <summary>
        /// Put the text on the clipboard - format is the clipboardData format ("Text" or "URL"); false if the host refuses it
        /// </summary>
        bool SetText(string format, string text);

        /// <summary>
        /// The text on the clipboard, or null if there is none in the format
        /// </summary>
        string? GetText(string format);

        /// <summary>
        /// Remove the format from the clipboard (all formats if format is empty); false if the host refuses it
        /// </summary>
        bool Clear(string format);
    }
}
