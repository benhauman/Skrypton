using System;
using System.Collections.Generic;
using System.Text;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    /// <summary>
    /// While active (on the current thread), this records which VBScript source line each line of the rendered C# program came from, so that a
    /// runtime error in the compiled program can be reported against the VBScript line (see issue #47, IScriptControl.Error.Line)
    /// </summary>
    internal sealed class SourceLineMapRecorder : IDisposable
    {
        [ThreadStatic]
        private static SourceLineMapRecorder? _current;

        private readonly SourceLineMapRecorder? _previous;
        private readonly SortedDictionary<int, int> _csLineToSourceLineIndex = new SortedDictionary<int, int>();
        private StringBuilder? _target;
        private int _scannedLength;
        private int _newLinesSeen;

        private SourceLineMapRecorder()
        {
            _previous = _current;
            _current = this;
        }

        internal static SourceLineMapRecorder? Current => _current;

        internal static SourceLineMapRecorder Start() => new SourceLineMapRecorder();

        /// <summary>
        /// Set the StringBuilder that the whole C# program is rendered into
        /// </summary>
        internal void SetTarget(StringBuilder sb)
        {
            _target = sb ?? throw new ArgumentNullException(nameof(sb));
            _scannedLength = 0;
            _newLinesSeen = 0;
            _csLineToSourceLineIndex.Clear();
        }

        /// <summary>
        /// Called before a builder renders itself into the StringBuilder. Builders are rendered parent-first, so when a parent and a nested builder
        /// start on the same C# line, the nested (more specific) builder's source line wins.
        /// </summary>
        internal void Record(StringBuilder sb, int sourceLineIndex)
        {
            if (sb == null)
                throw new ArgumentNullException(nameof(sb));

            // Only the program-level StringBuilder is tracked (builders may also render into temporary ones, whose lines are not program lines)
            if (!ReferenceEquals(sb, _target))
                return;
            for (int i = _scannedLength; i < sb.Length; i++)
            {
                if (sb[i] == '\n')
                    _newLinesSeen++;
            }
            _scannedLength = sb.Length;

            if (sourceLineIndex > 0)
                _csLineToSourceLineIndex[_newLinesSeen + 1] = sourceLineIndex; // C# lines are 1-based
        }

        /// <summary>
        /// The recorded map from 1-based C# line number to VBScript source line index
        /// </summary>
        internal IReadOnlyDictionary<int, int> CSharpLineToSourceLineIndex => _csLineToSourceLineIndex;

        public void Dispose()
        {
            if (ReferenceEquals(_current, this))
                _current = _previous;
        }
    }
}
