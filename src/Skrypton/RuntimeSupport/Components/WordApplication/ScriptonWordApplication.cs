// The members are bound by name from scripts, so they must be instance members
#pragma warning disable CA1822
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in Word.Application - Word documents (.docx) are opened, changed and saved without Word: the Word object model that scripts use
    /// to fill in documents (Documents.Open / Add, Tables, Rows, Cells, Bookmarks, Range.Text, SaveAs, Close) works on the WordprocessingML of the
    /// document. Files are read and written through the host's IHostFileSystemHostService (the process file system if there is none). There is
    /// no window: Visible and Caption are kept, nothing is shown.
    /// </summary>
    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Application")] // TypeName(CreateObject("Word.Application")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordApplication : ScriptonIReflectOnClrType
    {
        public const string ProgId = "Word.Application";
        internal const string ErrorSource = "Microsoft Word";

        private readonly ScriptonWordDocuments _documents;

        public ScriptonWordApplication(IRuntimeHost runtimeHost, Func<object?, string> toText) : base(typeof(ScriptonWordApplication))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            IHostFileSystemHostService fileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
            _documents = new ScriptonWordDocuments(fileSystem, toText ?? throw new ArgumentNullException(nameof(toText)));
        }

        /// <summary>
        /// The default member ("" &amp; CreateObject("Word.Application") is "Microsoft Word", VarType 8)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Name => "Microsoft Word";

        [DispId(8)] public string Caption { get; set; } = "Word";

        [DispId(23)] public bool Visible { get; set; }

        [DispId(6)] public object Documents => _documents;

        [DispId(3)] public object ActiveDocument => _documents.Active ?? throw new CustomException(4248, ErrorSource, "This command is not available because no document is open.");

        /// <summary>
        /// Quit([SaveChanges], ..) - the documents are closed (saved only if SaveChanges is True: there is nobody to ask)
        /// </summary>
        [DispId(1105)]
        public void Quit([In, Optional, DefaultParameterValue(null)] object? SaveChanges, [In, Optional, DefaultParameterValue(null)] object? OriginalFormat, [In, Optional, DefaultParameterValue(null)] object? RouteDocument)
        {
            foreach (ScriptonWordDocument document in _documents.All.ToList())
                document.Close(SaveChanges, null, null);
        }

        internal static CustomException MemberDoesNotExist() => new CustomException(5941, ErrorSource, "The requested member of the collection does not exist.");

        internal static int ToIndex(object? index) => Convert.ToInt32(index, CultureInfo.InvariantCulture);

        internal static bool IsMissing(object? value) => (value == null) || (value is Missing);
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Documents")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordDocuments : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly IHostFileSystemHostService _fileSystem;
        private readonly Func<object?, string> _toText;
        private readonly List<ScriptonWordDocument> _documents = new List<ScriptonWordDocument>();
        private int _newDocumentNumber;

        public ScriptonWordDocuments(IHostFileSystemHostService fileSystem, Func<object?, string> toText) : base(typeof(ScriptonWordDocuments))
        {
            _fileSystem = fileSystem;
            _toText = toText;
        }

        internal IReadOnlyList<ScriptonWordDocument> All => _documents;

        internal ScriptonWordDocument? Active => _documents.LastOrDefault();

        [DispId(2)] public int Count => _documents.Count;

        /// <summary>
        /// Item(Index) - the position (1-based) or the name of the document
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            if (Index is string name)
                return _documents.FirstOrDefault(d => string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase) || string.Equals(d.FullName, name, StringComparison.OrdinalIgnoreCase))
                    ?? throw ScriptonWordApplication.MemberDoesNotExist();
            int index = ScriptonWordApplication.ToIndex(Index);
            if ((index < 1) || (index > _documents.Count))
                throw ScriptonWordApplication.MemberDoesNotExist();
            return _documents[index - 1];
        }

        /// <summary>
        /// Add([Template], ..) - a new document, empty or a copy of the template
        /// </summary>
        [DispId(14)]
        public object Add([In, Optional, DefaultParameterValue(null)] object? Template, [In, Optional, DefaultParameterValue(null)] object? NewTemplate, [In, Optional, DefaultParameterValue(null)] object? DocumentType, [In, Optional, DefaultParameterValue(null)] object? Visible)
        {
            ScriptonWordprocessingDocument content = ScriptonWordApplication.IsMissing(Template) || (_toText(Template).Length == 0)
                ? ScriptonWordprocessingDocument.CreateEmpty()
                : LoadFile(_toText(Template)).Copy();
            _newDocumentNumber++;
            var document = new ScriptonWordDocument(this, _fileSystem, _toText, content, "Document" + _newDocumentNumber.ToString(CultureInfo.InvariantCulture), path: null);
            _documents.Add(document);
            return document;
        }

        /// <summary>
        /// Open(FileName, [ConfirmConversions], [ReadOnly], ..) - a document that is open already is returned as it is
        /// </summary>
        [DispId(19)]
        public object Open(
            [In] object FileName,
            [In, Optional, DefaultParameterValue(null)] object? ConfirmConversions,
            [In, Optional, DefaultParameterValue(null)] object? ReadOnly,
            [In, Optional, DefaultParameterValue(null)] object? AddToRecentFiles,
            [In, Optional, DefaultParameterValue(null)] object? PasswordDocument,
            [In, Optional, DefaultParameterValue(null)] object? PasswordTemplate,
            [In, Optional, DefaultParameterValue(null)] object? Revert,
            [In, Optional, DefaultParameterValue(null)] object? WritePasswordDocument,
            [In, Optional, DefaultParameterValue(null)] object? WritePasswordTemplate,
            [In, Optional, DefaultParameterValue(null)] object? Format,
            [In, Optional, DefaultParameterValue(null)] object? Encoding,
            [In, Optional, DefaultParameterValue(null)] object? Visible)
        {
            string path = _toText(FileName);
            ScriptonWordDocument? open = _documents.FirstOrDefault(d => string.Equals(d.FullName, path, StringComparison.OrdinalIgnoreCase));
            if (open != null)
                return open;
            var document = new ScriptonWordDocument(this, _fileSystem, _toText, LoadFile(path), ScriptonFileSystemObject.FileNameOf(path), path);
            _documents.Add(document);
            return document;
        }

        private ScriptonWordprocessingDocument LoadFile(string path)
        {
            if (!_fileSystem.FileExists(path))
                throw new CustomException(5174, ScriptonWordApplication.ErrorSource, "Sorry, we couldn't find your file. Was it moved, renamed, or deleted?\r(" + path + ")");
            try
            {
                return ScriptonWordprocessingDocument.Load(_fileSystem, path);
            }
            catch (Exception e) when ((e is System.IO.InvalidDataException) || (e is System.Xml.XmlException))
            {
                throw new CustomException(5792, ScriptonWordApplication.ErrorSource, "The file appears to be corrupted. (" + path + ")");
            }
        }

        internal void Closed(ScriptonWordDocument document) => _documents.Remove(document);

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _documents.ToList().GetEnumerator();
    }

    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Document")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordDocument : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocuments _documents;
        private readonly IHostFileSystemHostService _fileSystem;
        private readonly Func<object?, string> _toText;
        private string? _path;
        private bool _closed;

        public ScriptonWordDocument(ScriptonWordDocuments documents, IHostFileSystemHostService fileSystem, Func<object?, string> toText, ScriptonWordprocessingDocument content, string name, string? path)
            : base(typeof(ScriptonWordDocument))
        {
            _documents = documents;
            _fileSystem = fileSystem;
            _toText = toText;
            Content_ = content;
            Name = name;
            _path = path;
            Saved = true;
        }

        internal ScriptonWordprocessingDocument Content_ { get; }

        internal Func<object?, string> ToText => _toText;

        /// <summary>
        /// The default member: the file name ("Document1" for a new document that has not been saved)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Name { get; private set; }

        [DispId(9)] public string FullName => _path ?? Name;

        [DispId(5)] public string Path => (_path == null) ? "" : ScriptonFileSystemObject.ParentFolderNameOf(_path);

        [DispId(40)] public bool Saved { get; set; }

        [DispId(6)] public object Tables => new ScriptonWordTables(this);

        [DispId(4)] public object Bookmarks => new ScriptonWordBookmarks(this);

        [DispId(1019)] public object Content => new ScriptonWordRange(this, ScriptonWordRangeTarget.Content);

        [DispId(108)]
        public void Save()
        {
            EnsureOpen();
            if (_path == null)
                throw new NotSupportedException($"Document.Save: '{Name}' has not been saved yet - there is no Save As dialog, use SaveAs with a file name");
            Content_.Save(_fileSystem, _path);
            Saved = true;
        }

        /// <summary>
        /// SaveAs(FileName, [FileFormat], ..) - the document is saved as a Word document (.docx) whatever the format
        /// </summary>
        [DispId(376)]
        public void SaveAs([In] object FileName, [In, Optional, DefaultParameterValue(null)] object? FileFormat, [In, Optional, DefaultParameterValue(null)] object? LockComments, [In, Optional, DefaultParameterValue(null)] object? Password, [In, Optional, DefaultParameterValue(null)] object? AddToRecentFiles)
        {
            EnsureOpen();
            string path = _toText(FileName);
            if (path.Length == 0)
                throw new ArgumentException("Document.SaveAs: a file name is required", nameof(FileName));
            Content_.Save(_fileSystem, path);
            _path = path;
            Name = ScriptonFileSystemObject.FileNameOf(path);
            Saved = true;
        }

        [DispId(553)]
        public void SaveAs2([In] object FileName, [In, Optional, DefaultParameterValue(null)] object? FileFormat, [In, Optional, DefaultParameterValue(null)] object? LockComments, [In, Optional, DefaultParameterValue(null)] object? Password, [In, Optional, DefaultParameterValue(null)] object? AddToRecentFiles)
        {
            SaveAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles);
        }

        /// <summary>
        /// Close([SaveChanges], ..) - the changes are saved only if SaveChanges is True (wdSaveChanges, -1): there is nobody to ask
        /// </summary>
        [DispId(1105)]
        public void Close([In, Optional, DefaultParameterValue(null)] object? SaveChanges, [In, Optional, DefaultParameterValue(null)] object? OriginalFormat, [In, Optional, DefaultParameterValue(null)] object? RouteDocument)
        {
            if (_closed)
                return;
            if (!ScriptonWordApplication.IsMissing(SaveChanges) && (Convert.ToInt32(SaveChanges, CultureInfo.InvariantCulture) == -1) && !Saved)
                Save();
            _closed = true;
            _documents.Closed(this);
        }

        internal void Changed()
        {
            EnsureOpen();
            Saved = false;
        }

        internal void EnsureOpen()
        {
            if (_closed)
                throw new CustomException(5825, ScriptonWordApplication.ErrorSource, "Object has been deleted.");
        }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Tables")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordTables : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonWordDocument _document;

        public ScriptonWordTables(ScriptonWordDocument document) : base(typeof(ScriptonWordTables))
        {
            _document = document;
        }

        [DispId(2)] public int Count => _document.Content_.GetTables().Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            List<XElement> tables = _document.Content_.GetTables();
            int index = ScriptonWordApplication.ToIndex(Index);
            if ((index < 1) || (index > tables.Count))
                throw ScriptonWordApplication.MemberDoesNotExist();
            return new ScriptonWordTable(_document, tables[index - 1]);
        }

        /// <summary>
        /// Add(Range, NumRows, NumColumns, ..) - the range must be Document.Content (the table replaces the content)
        /// </summary>
        [DispId(200)]
        public object Add([In] object Range, [In] object NumRows, [In] object NumColumns, [In, Optional, DefaultParameterValue(null)] object? DefaultTableBehavior, [In, Optional, DefaultParameterValue(null)] object? AutoFitBehavior)
        {
            if (!(Range is ScriptonWordRange range) || !range.Target.IsContent)
                throw new NotSupportedException("Tables.Add: only Document.Content is supported as the range (the table replaces the content)");
            int rows = ScriptonWordApplication.ToIndex(NumRows);
            int columns = ScriptonWordApplication.ToIndex(NumColumns);
            if ((rows < 1) || (columns < 1) || (columns > 63))
                throw new CustomException(5149, ScriptonWordApplication.ErrorSource, "The measurement must be between 1 and 63.");
            _document.Changed();
            return new ScriptonWordTable(_document, _document.Content_.ReplaceContentWithTable(rows, columns));
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _document.Content_.GetTables().Select(t => new ScriptonWordTable(_document, t)).ToList().GetEnumerator();
    }

    [SourceClassName("Table")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordTable : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocument _document;
        private readonly XElement _table;

        public ScriptonWordTable(ScriptonWordDocument document, XElement table) : base(typeof(ScriptonWordTable))
        {
            _document = document;
            _table = table;
        }

        [DispId(101)] public object Rows => new ScriptonWordRows(_document, _table);

        /// <summary>
        /// Cell(Row, Column) - 1-based
        /// </summary>
        [DispId(17)]
        public object Cell([In] int Row, [In] int Column)
        {
            XElement row = _table.Elements(ScriptonWordprocessingDocument.W + "tr").ElementAtOrDefault(Row - 1) ?? throw ScriptonWordApplication.MemberDoesNotExist();
            XElement cell = row.Elements(ScriptonWordprocessingDocument.W + "tc").ElementAtOrDefault(Column - 1) ?? throw ScriptonWordApplication.MemberDoesNotExist();
            return new ScriptonWordCell(_document, cell);
        }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Rows")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordRows : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonWordDocument _document;
        private readonly XElement _table;

        public ScriptonWordRows(ScriptonWordDocument document, XElement table) : base(typeof(ScriptonWordRows))
        {
            _document = document;
            _table = table;
        }

        private List<XElement> All => _table.Elements(ScriptonWordprocessingDocument.W + "tr").ToList();

        [DispId(2)] public int Count => All.Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] int Index)
        {
            List<XElement> rows = All;
            if ((Index < 1) || (Index > rows.Count))
                throw ScriptonWordApplication.MemberDoesNotExist();
            return new ScriptonWordRow(_document, rows[Index - 1]);
        }

        /// <summary>
        /// Add([BeforeRow]) - a row like the row it is added before (or the last row), without its text
        /// </summary>
        [DispId(100)]
        public object Add([In, Optional, DefaultParameterValue(null)] object? BeforeRow)
        {
            XElement? before = null;
            if (!ScriptonWordApplication.IsMissing(BeforeRow))
                before = (BeforeRow as ScriptonWordRow)?.Element ?? throw new ArgumentException("Rows.Add: BeforeRow must be a Row of the table", nameof(BeforeRow));
            _document.Changed();
            return new ScriptonWordRow(_document, ScriptonWordprocessingDocument.AddRow(_table, before));
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => All.Select(r => new ScriptonWordRow(_document, r)).ToList().GetEnumerator();
    }

    [SourceClassName("Row")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordRow : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocument _document;

        public ScriptonWordRow(ScriptonWordDocument document, XElement row) : base(typeof(ScriptonWordRow))
        {
            _document = document;
            Element = row;
        }

        internal XElement Element { get; }

        [DispId(100)] public object Cells => new ScriptonWordCells(_document, Element);

        [DispId(4)] public int Index => Element.ElementsBeforeSelf(ScriptonWordprocessingDocument.W + "tr").Count() + 1;
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Cells")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordCells : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonWordDocument _document;
        private readonly XElement _row;

        public ScriptonWordCells(ScriptonWordDocument document, XElement row) : base(typeof(ScriptonWordCells))
        {
            _document = document;
            _row = row;
        }

        private List<XElement> All => _row.Elements(ScriptonWordprocessingDocument.W + "tc").ToList();

        [DispId(2)] public int Count => All.Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] int Index)
        {
            List<XElement> cells = All;
            if ((Index < 1) || (Index > cells.Count))
                throw ScriptonWordApplication.MemberDoesNotExist();
            return new ScriptonWordCell(_document, cells[Index - 1]);
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => All.Select(c => new ScriptonWordCell(_document, c)).ToList().GetEnumerator();
    }

    /// <summary>
    /// A table cell - it has no default member (VarType 9); cell.Range = "text" sets the text of its range (as with Word)
    /// </summary>
    [SourceClassName("Cell")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordCell : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocument _document;
        private readonly XElement _cell;

        public ScriptonWordCell(ScriptonWordDocument document, XElement cell) : base(typeof(ScriptonWordCell))
        {
            _document = document;
            _cell = cell;
        }

        [DispId(1)]
        public object Range
        {
            get => new ScriptonWordRange(_document, ScriptonWordRangeTarget.ForCell(_cell));
            set => new ScriptonWordRange(_document, ScriptonWordRangeTarget.ForCell(_cell)).Text = value;
        }

        [DispId(4)] public int RowIndex => _cell.Parent!.ElementsBeforeSelf(ScriptonWordprocessingDocument.W + "tr").Count() + 1;

        [DispId(5)] public int ColumnIndex => _cell.ElementsBeforeSelf(ScriptonWordprocessingDocument.W + "tc").Count() + 1;
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Bookmarks")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordBookmarks : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonWordDocument _document;

        public ScriptonWordBookmarks(ScriptonWordDocument document) : base(typeof(ScriptonWordBookmarks))
        {
            _document = document;
        }

        [DispId(2)] public int Count => _document.Content_.GetBookmarkNames().Count;

        /// <summary>
        /// Item(Index) - the position (1-based, in the order of the names) or the name (case does not matter)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            string name;
            if (Index is string text)
            {
                name = text;
            }
            else
            {
                List<string> names = _document.Content_.GetBookmarkNames();
                int index = ScriptonWordApplication.ToIndex(Index);
                if ((index < 1) || (index > names.Count))
                    throw ScriptonWordApplication.MemberDoesNotExist();
                name = names[index - 1];
            }
            XElement start = _document.Content_.FindBookmarkStart(name) ?? throw ScriptonWordApplication.MemberDoesNotExist();
            return new ScriptonWordBookmark(_document, (string?)start.Attribute(ScriptonWordprocessingDocument.W + "name") ?? name);
        }

        [DispId(6)]
        public bool Exists([In, MarshalAs(UnmanagedType.BStr)] string Name) => _document.Content_.FindBookmarkStart(Name ?? "") != null;

        /// <summary>
        /// Add(Name, [Range]) - the range is a cell, a bookmark or the whole document (the default)
        /// </summary>
        [DispId(5)]
        public object Add([In, MarshalAs(UnmanagedType.BStr)] string Name, [In, Optional, DefaultParameterValue(null)] object? Range)
        {
            if (string.IsNullOrEmpty(Name))
                throw new CustomException(5828, ScriptonWordApplication.ErrorSource, "Bad bookmark name.");
            ScriptonWordRangeTarget target = ScriptonWordRangeTarget.Content;
            if (!ScriptonWordApplication.IsMissing(Range))
                target = (Range as ScriptonWordRange)?.Target ?? throw new ArgumentException("Bookmarks.Add: Range must be a Range", nameof(Range));
            _document.Changed();
            _document.Content_.AddBookmark(Name, target);
            return new ScriptonWordBookmark(_document, Name);
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _document.Content_.GetBookmarkNames().Select(n => new ScriptonWordBookmark(_document, n)).ToList().GetEnumerator();
    }

    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Bookmark")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordBookmark : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocument _document;

        public ScriptonWordBookmark(ScriptonWordDocument document, string name) : base(typeof(ScriptonWordBookmark))
        {
            _document = document;
            Name = name;
        }

        [DispId(0)]
        [IsDefault]
        public string Name { get; }

        [DispId(1)] public object Range => new ScriptonWordRange(_document, ScriptonWordRangeTarget.ForBookmark(Start()));

        [DispId(11)]
        public void Delete()
        {
            _document.Changed();
            _document.Content_.RemoveBookmark(Start());
        }

        private XElement Start() => _document.Content_.FindBookmarkStart(Name) ?? throw ScriptonWordApplication.MemberDoesNotExist();
    }

    /// <summary>
    /// A range: the whole document (Document.Content), a cell or a bookmark - its default member is Text
    /// </summary>
    [DefaultMember("Text")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Range")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWordRange : ScriptonIReflectOnClrType
    {
        private readonly ScriptonWordDocument _document;

        public ScriptonWordRange(ScriptonWordDocument document, ScriptonWordRangeTarget target) : base(typeof(ScriptonWordRange))
        {
            _document = document;
            Target = target;
        }

        internal ScriptonWordRangeTarget Target { get; }

        [DispId(0)]
        [IsDefault]
        public object Text
        {
            get
            {
                _document.EnsureOpen();
                if (Target.Cell != null)
                    return ScriptonWordprocessingDocument.GetCellText(Target.Cell);
                if (Target.BookmarkStart != null)
                    return _document.Content_.GetBookmarkText(Target.BookmarkStart);
                return _document.Content_.GetContentText();
            }
            set
            {
                if ((value == null) || (value is DBNull))
                    throw new CustomException(13, ScriptonWordApplication.ErrorSource, "Type mismatch");
                string text = _document.ToText(value);
                _document.Changed();
                if (Target.Cell != null)
                    ScriptonWordprocessingDocument.SetCellText(Target.Cell, text);
                else if (Target.BookmarkStart != null)
                    _document.Content_.SetBookmarkText(Target.BookmarkStart, text);
                else
                    _document.Content_.SetContentText(text);
            }
        }
    }
}
