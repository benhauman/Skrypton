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
    /// The built-in Excel.Application - Excel workbooks (.xlsx) are created, opened, changed and saved without Excel: the Excel object model that
    /// scripts use for reports (Workbooks.Add / Open, Worksheets, Cells, Range, Value, Formula, NumberFormat, Font.Bold, SaveAs, Close) works
    /// on the SpreadsheetML of the workbook. Files are read and written through the host's IHostFileSystemHostService (the process file system
    /// if there is none). There is no window and no calculation: Visible, DisplayAlerts and ScreenUpdating are kept, formulas are calculated by
    /// Excel when the workbook is opened, AutoFit does nothing.
    /// </summary>
    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Application")] // TypeName(CreateObject("Excel.Application")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelApplication : ScriptonIReflectOnClrType
    {
        public const string ProgId = "Excel.Application";
        internal const string ErrorSource = "Microsoft Excel";

        private readonly ScriptonExcelWorkbooks _workbooks;

        public ScriptonExcelApplication(IRuntimeHost runtimeHost, Func<object?, string> toText) : base(typeof(ScriptonExcelApplication))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            IHostFileSystemHostService fileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
            _workbooks = new ScriptonExcelWorkbooks(fileSystem, toText ?? throw new ArgumentNullException(nameof(toText)));
        }

        /// <summary>
        /// The default member ("" &amp; CreateObject("Excel.Application") is "Microsoft Excel", VarType 8)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Name => "Microsoft Excel";

        [DispId(392)] public string Version => "16.0";

        [DispId(558)] public bool Visible { get; set; }

        [DispId(343)] public bool DisplayAlerts { get; set; } = true;

        [DispId(382)] public bool ScreenUpdating { get; set; } = true;

        [DispId(572)] public object Workbooks => _workbooks;

        /// <summary>
        /// The active workbook - Nothing if there is none (as with Excel)
        /// </summary>
        [DispId(308)] public object ActiveWorkbook => (object?)_workbooks.Active ?? VBScriptConstants.Nothing;

        [DispId(307)] public object ActiveSheet => (object?)_workbooks.Active?.ActiveWorksheet ?? VBScriptConstants.Nothing;

        [DispId(238)]
        public object Cells([In, Optional, DefaultParameterValue(null)] object? RowIndex, [In, Optional, DefaultParameterValue(null)] object? ColumnIndex) => RequireActiveSheet().Cells(RowIndex, ColumnIndex);

        [DispId(197)]
        public object Range([In] object Cell1, [In, Optional, DefaultParameterValue(null)] object? Cell2) => RequireActiveSheet().Range(Cell1, Cell2);

        [DispId(241)]
        public object Columns([In, Optional, DefaultParameterValue(null)] object? Index) => RequireActiveSheet().Columns(Index);

        [DispId(258)]
        public object Rows([In, Optional, DefaultParameterValue(null)] object? Index) => RequireActiveSheet().Rows(Index);

        /// <summary>
        /// Quit - the workbooks are closed without saving (there is nobody to ask)
        /// </summary>
        [DispId(302)]
        public void Quit()
        {
            foreach (ScriptonExcelWorkbook workbook in _workbooks.All.ToList())
                workbook.Close(false, null, null);
        }

        private ScriptonExcelWorksheet RequireActiveSheet() => _workbooks.Active?.ActiveWorksheet ?? throw new CustomException(91, ErrorSource, "Object variable not set");

        internal static CustomException SubscriptOutOfRange() => new CustomException(9, ErrorSource, "Subscript out of range");

        internal static bool IsMissing(object? value) => (value == null) || (value is Missing);

        internal static int ToInt(object? value) => Convert.ToInt32(value, CultureInfo.InvariantCulture);

        /// <summary>
        /// A value written to a cell the way Excel takes it: numbers become Double, text that is a number or TRUE / FALSE becomes that
        /// </summary>
        internal static object? ToCellValue(object? value)
        {
            switch (value)
            {
                case null:
                case DBNull _:
                    return null;
                case bool _:
                case DateTime _:
                case double _:
                    return value;
                case string text:
                    if (text.StartsWith("=", StringComparison.Ordinal))
                        return text;
                    if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double number))
                        return number;
                    if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
                        return false;
                    return text;
                case IConvertible convertible:
                    return convertible.ToDouble(CultureInfo.InvariantCulture);
                default:
                    throw new CustomException(13, ErrorSource, "Type mismatch");
            }
        }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Workbooks")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelWorkbooks : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly IHostFileSystemHostService _fileSystem;
        private readonly Func<object?, string> _toText;
        private readonly List<ScriptonExcelWorkbook> _workbooks = new List<ScriptonExcelWorkbook>();
        private int _newWorkbookNumber;

        public ScriptonExcelWorkbooks(IHostFileSystemHostService fileSystem, Func<object?, string> toText) : base(typeof(ScriptonExcelWorkbooks))
        {
            _fileSystem = fileSystem;
            _toText = toText;
        }

        internal IReadOnlyList<ScriptonExcelWorkbook> All => _workbooks;

        internal ScriptonExcelWorkbook? Active { get; set; }

        [DispId(118)] public int Count => _workbooks.Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            if (Index is string name)
                return _workbooks.FirstOrDefault(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase)) ?? throw ScriptonExcelApplication.SubscriptOutOfRange();
            int index = ScriptonExcelApplication.ToInt(Index);
            if ((index < 1) || (index > _workbooks.Count))
                throw ScriptonExcelApplication.SubscriptOutOfRange();
            return _workbooks[index - 1];
        }

        /// <summary>
        /// Add([Template]) - a new workbook with one sheet, or a copy of the template
        /// </summary>
        [DispId(181)]
        public object Add([In, Optional, DefaultParameterValue(null)] object? Template)
        {
            ScriptonSpreadsheetDocument content = (ScriptonExcelApplication.IsMissing(Template) || (Template is int) || (Template is short))
                ? ScriptonSpreadsheetDocument.CreateEmpty()
                : LoadFile(_toText(Template)).Copy();
            _newWorkbookNumber++;
            return Opened(new ScriptonExcelWorkbook(this, _fileSystem, _toText, content, "Book" + _newWorkbookNumber.ToString(CultureInfo.InvariantCulture), path: null));
        }

        /// <summary>
        /// Open(Filename, ..) - a workbook that is open already is returned as it is
        /// </summary>
        [DispId(1923)]
        public object Open(
            [In] object Filename,
            [In, Optional, DefaultParameterValue(null)] object? UpdateLinks,
            [In, Optional, DefaultParameterValue(null)] object? ReadOnly,
            [In, Optional, DefaultParameterValue(null)] object? Format,
            [In, Optional, DefaultParameterValue(null)] object? Password,
            [In, Optional, DefaultParameterValue(null)] object? WriteResPassword,
            [In, Optional, DefaultParameterValue(null)] object? IgnoreReadOnlyRecommended,
            [In, Optional, DefaultParameterValue(null)] object? Origin,
            [In, Optional, DefaultParameterValue(null)] object? Delimiter,
            [In, Optional, DefaultParameterValue(null)] object? Editable,
            [In, Optional, DefaultParameterValue(null)] object? Notify,
            [In, Optional, DefaultParameterValue(null)] object? Converter,
            [In, Optional, DefaultParameterValue(null)] object? AddToMru)
        {
            string path = _toText(Filename);
            ScriptonExcelWorkbook? open = _workbooks.FirstOrDefault(w => string.Equals(w.FullName, path, StringComparison.OrdinalIgnoreCase));
            if (open != null)
                return open;
            return Opened(new ScriptonExcelWorkbook(this, _fileSystem, _toText, LoadFile(path), ScriptonFileSystemObject.FileNameOf(path), path));
        }

        private ScriptonExcelWorkbook Opened(ScriptonExcelWorkbook workbook)
        {
            _workbooks.Add(workbook);
            Active = workbook;
            return workbook;
        }

        private ScriptonSpreadsheetDocument LoadFile(string path)
        {
            if (!_fileSystem.FileExists(path))
                throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "Sorry, we couldn't find " + path + ". Is it possible it was moved, renamed or deleted?");
            try
            {
                return ScriptonSpreadsheetDocument.Load(_fileSystem, path);
            }
            catch (Exception e) when ((e is System.IO.InvalidDataException) || (e is System.Xml.XmlException))
            {
                throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "Excel cannot open the file '" + path + "' because the file format or file extension is not valid.");
            }
        }

        internal void Closed(ScriptonExcelWorkbook workbook)
        {
            _workbooks.Remove(workbook);
            if (Active == workbook)
                Active = _workbooks.LastOrDefault();
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _workbooks.ToList().GetEnumerator();
    }

    /// <summary>
    /// A workbook - it has no default member (VarType 9)
    /// </summary>
    [SourceClassName("Workbook")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelWorkbook : ScriptonIReflectOnClrType
    {
        private readonly ScriptonExcelWorkbooks _workbooks;
        private readonly IHostFileSystemHostService _fileSystem;
        private string? _path;
        private bool _closed;

        public ScriptonExcelWorkbook(ScriptonExcelWorkbooks workbooks, IHostFileSystemHostService fileSystem, Func<object?, string> toText, ScriptonSpreadsheetDocument content, string name, string? path)
            : base(typeof(ScriptonExcelWorkbook))
        {
            _workbooks = workbooks;
            _fileSystem = fileSystem;
            ToText = toText;
            Content = content;
            Name = name;
            _path = path;
            Saved = true;
        }

        internal ScriptonSpreadsheetDocument Content { get; }

        internal Func<object?, string> ToText { get; }

        internal ScriptonExcelWorksheet ActiveWorksheet
        {
            get
            {
                List<XElement> sheets = Content.GetSheetElements();
                return new ScriptonExcelWorksheet(this, sheets[Math.Min(Math.Max(Content.ActiveSheetIndex, 0), sheets.Count - 1)]);
            }
        }

        [DispId(110)] public string Name { get; private set; }

        [DispId(289)] public string FullName => _path ?? Name;

        [DispId(291)] public string Path => (_path == null) ? "" : ScriptonFileSystemObject.ParentFolderNameOf(_path);

        [DispId(298)] public bool Saved { get; set; }

        [DispId(494)] public object Worksheets => new ScriptonExcelSheets(this);

        [DispId(485)] public object Sheets => new ScriptonExcelSheets(this);

        [DispId(307)] public object ActiveSheet => ActiveWorksheet;

        [DispId(304)]
        public void Activate()
        {
            EnsureOpen();
            _workbooks.Active = this;
        }

        [DispId(283)]
        public void Save()
        {
            EnsureOpen();
            if (_path == null)
                throw new NotSupportedException($"Workbook.Save: '{Name}' has not been saved yet - there is no Save As dialog, use SaveAs with a file name");
            Content.Save(_fileSystem, _path);
            Saved = true;
        }

        /// <summary>
        /// SaveAs(Filename, [FileFormat], ..) - the workbook is saved as an Excel workbook (.xlsx) whatever the format
        /// </summary>
        [DispId(1925)]
        public void SaveAs(
            [In] object Filename,
            [In, Optional, DefaultParameterValue(null)] object? FileFormat,
            [In, Optional, DefaultParameterValue(null)] object? Password,
            [In, Optional, DefaultParameterValue(null)] object? WriteResPassword,
            [In, Optional, DefaultParameterValue(null)] object? ReadOnlyRecommended,
            [In, Optional, DefaultParameterValue(null)] object? CreateBackup)
        {
            EnsureOpen();
            string path = ToText(Filename);
            if (path.Length == 0)
                throw new ArgumentException("Workbook.SaveAs: a file name is required", nameof(Filename));
            Content.Save(_fileSystem, path);
            _path = path;
            Name = ScriptonFileSystemObject.FileNameOf(path);
            Saved = true;
        }

        /// <summary>
        /// Close([SaveChanges], [Filename], ..) - the changes are saved only if SaveChanges is True: there is nobody to ask
        /// </summary>
        [DispId(277)]
        public void Close([In, Optional, DefaultParameterValue(null)] object? SaveChanges, [In, Optional, DefaultParameterValue(null)] object? Filename, [In, Optional, DefaultParameterValue(null)] object? RouteWorkbook)
        {
            if (_closed)
                return;
            if (!ScriptonExcelApplication.IsMissing(SaveChanges) && Convert.ToBoolean(SaveChanges, CultureInfo.InvariantCulture) && !Saved)
            {
                if (!ScriptonExcelApplication.IsMissing(Filename))
                    SaveAs(Filename!);
                else
                    Save();
            }
            _closed = true;
            _workbooks.Closed(this);
        }

        internal void Changed()
        {
            EnsureOpen();
            Saved = false;
        }

        internal void EnsureOpen()
        {
            if (_closed)
                throw new CustomException(-2147221080, ScriptonExcelApplication.ErrorSource, "Object has been deleted (the workbook is closed)");
        }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Sheets")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelSheets : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonExcelWorkbook _workbook;

        public ScriptonExcelSheets(ScriptonExcelWorkbook workbook) : base(typeof(ScriptonExcelSheets))
        {
            _workbook = workbook;
        }

        [DispId(118)] public int Count => _workbook.Content.GetSheetElements().Count;

        /// <summary>
        /// Item(Index) - the position (1-based) or the name of the sheet
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            List<XElement> sheets = _workbook.Content.GetSheetElements();
            XElement? sheet;
            if (Index is string name)
            {
                sheet = sheets.FirstOrDefault(s => string.Equals(ScriptonSpreadsheetDocument.GetSheetName(s), name, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                int index = ScriptonExcelApplication.ToInt(Index);
                sheet = ((index >= 1) && (index <= sheets.Count)) ? sheets[index - 1] : null;
            }
            return new ScriptonExcelWorksheet(_workbook, sheet ?? throw ScriptonExcelApplication.SubscriptOutOfRange());
        }

        /// <summary>
        /// Add([Before], [After], ..) - a new sheet before the active sheet (or before / after the given sheet); it becomes the active sheet
        /// </summary>
        [DispId(181)]
        public object Add([In, Optional, DefaultParameterValue(null)] object? Before, [In, Optional, DefaultParameterValue(null)] object? After, [In, Optional, DefaultParameterValue(null)] object? Count, [In, Optional, DefaultParameterValue(null)] object? Type)
        {
            _workbook.Changed();
            List<XElement> sheets = _workbook.Content.GetSheetElements();
            int position = _workbook.Content.ActiveSheetIndex;
            if (Before is ScriptonExcelWorksheet before)
                position = sheets.IndexOf(before.Element);
            else if (After is ScriptonExcelWorksheet after)
                position = sheets.IndexOf(after.Element) + 1;
            XElement sheet = _workbook.Content.AddSheet(position);
            _workbook.Content.ActiveSheetIndex = _workbook.Content.GetSheetElements().IndexOf(sheet);
            return new ScriptonExcelWorksheet(_workbook, sheet);
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _workbook.Content.GetSheetElements().Select(s => new ScriptonExcelWorksheet(_workbook, s)).ToList().GetEnumerator();
    }

    /// <summary>
    /// A worksheet - it has no default member (VarType 9)
    /// </summary>
    [SourceClassName("Worksheet")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelWorksheet : ScriptonIReflectOnClrType
    {
        public ScriptonExcelWorksheet(ScriptonExcelWorkbook workbook, XElement sheet) : base(typeof(ScriptonExcelWorksheet))
        {
            Workbook = workbook;
            Element = sheet;
        }

        private static readonly char[] InvalidSheetNameCharacters = { ':', '\\', '/', '?', '*', '[', ']' };

        internal ScriptonExcelWorkbook Workbook { get; }

        internal XElement Element { get; }

        internal XElement SheetData => Workbook.Content.GetSheetData(Element);

        [DispId(110)]
        public string Name
        {
            get => ScriptonSpreadsheetDocument.GetSheetName(Element);
            set
            {
                if (string.IsNullOrEmpty(value) || (value.Length > 31) || (value.IndexOfAny(InvalidSheetNameCharacters) >= 0))
                    throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "You typed an invalid name for a sheet or chart.");
                if (Workbook.Content.GetSheetElements().Any(s => (s != Element) && string.Equals(ScriptonSpreadsheetDocument.GetSheetName(s), value, StringComparison.OrdinalIgnoreCase)))
                    throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "That name is already taken. Try a different one.");
                Workbook.Changed();
                ScriptonSpreadsheetDocument.SetSheetName(Element, value);
            }
        }

        [DispId(486)] public int Index => Workbook.Content.GetSheetElements().IndexOf(Element) + 1;

        /// <summary>
        /// Cells([RowIndex], [ColumnIndex]) - the cell (the column is a number or letters), or all cells
        /// </summary>
        [DispId(238)]
        public object Cells([In, Optional, DefaultParameterValue(null)] object? RowIndex, [In, Optional, DefaultParameterValue(null)] object? ColumnIndex)
        {
            if (ScriptonExcelApplication.IsMissing(RowIndex))
                return new ScriptonExcelRange(this, 1, 1, 1048576, 16384);
            int row = ScriptonExcelApplication.ToInt(RowIndex);
            int column = (ColumnIndex is string letters) ? ScriptonSpreadsheetDocument.ColumnIndex(letters) : ScriptonExcelApplication.ToInt(ColumnIndex);
            if ((row < 1) || (column < 1))
                throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "Application-defined or object-defined error");
            return new ScriptonExcelRange(this, row, column, row, column);
        }

        /// <summary>
        /// Range(Cell1, [Cell2]) - an address ("A1", "A1:C3") or the cells between two ranges
        /// </summary>
        [DispId(197)]
        public object Range([In] object Cell1, [In, Optional, DefaultParameterValue(null)] object? Cell2)
        {
            ScriptonExcelRange first = ToRange(Cell1);
            if (ScriptonExcelApplication.IsMissing(Cell2))
                return first;
            ScriptonExcelRange second = ToRange(Cell2!);
            return new ScriptonExcelRange(this, Math.Min(first.FirstRow, second.FirstRow), Math.Min(first.FirstColumn, second.FirstColumn), Math.Max(first.LastRow, second.LastRow), Math.Max(first.LastColumn, second.LastColumn));
        }

        private ScriptonExcelRange ToRange(object cell)
        {
            if (cell is ScriptonExcelRange range)
                return range;
            string address = Workbook.ToText(cell);
            try
            {
                string[] parts = address.Split(':');
                (int row1, int column1) = ScriptonSpreadsheetDocument.ParseAddress(parts[0]);
                (int row2, int column2) = (parts.Length > 1) ? ScriptonSpreadsheetDocument.ParseAddress(parts[1]) : (row1, column1);
                if (parts.Length > 2)
                    throw new FormatException(address);
                return new ScriptonExcelRange(this, Math.Min(row1, row2), Math.Min(column1, column2), Math.Max(row1, row2), Math.Max(column1, column2));
            }
            catch (FormatException)
            {
                throw new CustomException(1004, ScriptonExcelApplication.ErrorSource, "Method 'Range' of object '_Worksheet' failed");
            }
        }

        /// <summary>
        /// Columns([Index]) - all columns, a column (number or letters) or columns ("A:C")
        /// </summary>
        [DispId(241)]
        public object Columns([In, Optional, DefaultParameterValue(null)] object? Index)
        {
            if (ScriptonExcelApplication.IsMissing(Index))
                return new ScriptonExcelRange(this, 1, 1, 1048576, 16384, ScriptonExcelRange.Axis.Columns);
            if (Index is string letters)
            {
                string[] parts = letters.Replace("$", "").Split(':');
                int first = ScriptonSpreadsheetDocument.ColumnIndex(parts[0]);
                int last = (parts.Length > 1) ? ScriptonSpreadsheetDocument.ColumnIndex(parts[1]) : first;
                return new ScriptonExcelRange(this, 1, Math.Min(first, last), 1048576, Math.Max(first, last), ScriptonExcelRange.Axis.Columns);
            }
            int column = ScriptonExcelApplication.ToInt(Index);
            return new ScriptonExcelRange(this, 1, column, 1048576, column, ScriptonExcelRange.Axis.Columns);
        }

        /// <summary>
        /// Rows([Index]) - all rows, a row or rows ("1:3")
        /// </summary>
        [DispId(258)]
        public object Rows([In, Optional, DefaultParameterValue(null)] object? Index)
        {
            if (ScriptonExcelApplication.IsMissing(Index))
                return new ScriptonExcelRange(this, 1, 1, 1048576, 16384, ScriptonExcelRange.Axis.Rows);
            if (Index is string text)
            {
                string[] parts = text.Split(':');
                int first = int.Parse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture);
                int last = (parts.Length > 1) ? int.Parse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture) : first;
                return new ScriptonExcelRange(this, Math.Min(first, last), 1, Math.Max(first, last), 16384, ScriptonExcelRange.Axis.Rows);
            }
            int row = ScriptonExcelApplication.ToInt(Index);
            return new ScriptonExcelRange(this, row, 1, row, 16384, ScriptonExcelRange.Axis.Rows);
        }

        /// <summary>
        /// The cells from the first to the last used cell
        /// </summary>
        [DispId(412)]
        public object UsedRange
        {
            get
            {
                var used = ScriptonSpreadsheetDocument.GetUsedRange(SheetData);
                return new ScriptonExcelRange(this, used.FirstRow, used.FirstColumn, used.LastRow, used.LastColumn);
            }
        }

        [DispId(304)]
        public void Activate()
        {
            Workbook.Activate();
            Workbook.Content.ActiveSheetIndex = Index - 1;
        }

        [DispId(235)]
        public void Select([In, Optional, DefaultParameterValue(null)] object? Replace) => Activate();
    }

    /// <summary>
    /// A range of cells - its default member is Value (VarType of a cell is the VarType of its value)
    /// </summary>
    [DefaultMember("Value")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Range")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelRange : ScriptonIReflectOnClrType
    {
        internal enum Axis
        {
            Cells,
            Rows,
            Columns,
        }

        private static readonly int[] OneBased = { 1, 1 }; // the arrays of values are 1-based, as with Excel

        private readonly ScriptonExcelWorksheet _sheet;
        private readonly Axis _axis;

        public ScriptonExcelRange(ScriptonExcelWorksheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn, Axis axis = Axis.Cells) : base(typeof(ScriptonExcelRange))
        {
            _sheet = sheet;
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            _axis = axis;
        }

        internal int FirstRow { get; }
        internal int FirstColumn { get; }
        internal int LastRow { get; }
        internal int LastColumn { get; }

        private ScriptonSpreadsheetDocument Content => _sheet.Workbook.Content;

        private IEnumerable<(int Row, int Column)> AllCells()
        {
            if (((long)(LastRow - FirstRow + 1) * (LastColumn - FirstColumn + 1)) > 1000000)
                throw new NotSupportedException("Range: changing more than 1,000,000 cells at once is not supported");
            for (int row = FirstRow; row <= LastRow; row++)
                for (int column = FirstColumn; column <= LastColumn; column++)
                    yield return (row, column);
        }

        /// <summary>
        /// The value of a cell (Empty, Double, String, Boolean, Date), or a two-dimensional array (1-based) of the values of more cells - setting it
        /// sets every cell (or the cells from a two-dimensional array)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object? Value
        {
            get
            {
                XElement sheetData = _sheet.SheetData;
                if ((FirstRow == LastRow) && (FirstColumn == LastColumn))
                    return Content.GetValue(sheetData, FirstRow, FirstColumn);
                var values = Array.CreateInstance(typeof(object), new[] { LastRow - FirstRow + 1, LastColumn - FirstColumn + 1 }, OneBased);
                foreach (var (row, column) in AllCells())
                    values.SetValue(Content.GetValue(sheetData, row, column), row - FirstRow + 1, column - FirstColumn + 1);
                return values;
            }
            set
            {
                _sheet.Workbook.Changed();
                XElement sheetData = _sheet.SheetData;
                if ((value is Array array) && (array.Rank == 2))
                {
                    int rowBase = array.GetLowerBound(0);
                    int columnBase = array.GetLowerBound(1);
                    foreach (var (row, column) in AllCells())
                    {
                        int i = row - FirstRow + rowBase;
                        int j = column - FirstColumn + columnBase;
                        object? item = ((i <= array.GetUpperBound(0)) && (j <= array.GetUpperBound(1))) ? array.GetValue(i, j) : null;
                        Content.SetValue(sheetData, row, column, ScriptonExcelApplication.ToCellValue(item));
                    }
                    return;
                }
                object? cellValue = ScriptonExcelApplication.ToCellValue(value);
                foreach (var (row, column) in AllCells())
                    Content.SetValue(sheetData, row, column, cellValue);
            }
        }

        [DispId(1388)]
        public object? Value2
        {
            get
            {
                object? value = Value;
                return (value is DateTime date) ? date.ToOADate() : value;
            }
            set => Value = value;
        }

        /// <summary>
        /// The text of the (first) cell - without Excel's column width rules ("####")
        /// </summary>
        [DispId(138)]
        public string Text
        {
            get
            {
                object? value = Content.GetValue(_sheet.SheetData, FirstRow, FirstColumn);
                switch (value)
                {
                    case null:
                        return "";
                    case bool boolean:
                        return boolean ? "TRUE" : "FALSE";
                    case DateTime date:
                        return date.ToString((date.TimeOfDay == TimeSpan.Zero) ? "d" : "g", CultureInfo.CurrentCulture);
                    default:
                        return _sheet.Workbook.ToText(value);
                }
            }
        }

        /// <summary>
        /// The formula of the (first) cell ("=..") or its value as text; setting it sets the cells
        /// </summary>
        [DispId(261)]
        public object Formula
        {
            get
            {
                XElement sheetData = _sheet.SheetData;
                string? formula = ScriptonSpreadsheetDocument.GetFormula(sheetData, FirstRow, FirstColumn);
                if (formula != null)
                    return "=" + formula;
                object? value = Content.GetValue(sheetData, FirstRow, FirstColumn);
                return value switch
                {
                    null => "",
                    bool boolean => boolean ? "TRUE" : "FALSE",
                    DateTime date => date.ToOADate().ToString(CultureInfo.InvariantCulture),
                    double number => number.ToString(CultureInfo.InvariantCulture),
                    _ => (string)value,
                };
            }
            set => Value = value;
        }

        [DispId(193)]
        public object NumberFormat
        {
            get => Content.GetNumberFormat(_sheet.SheetData, FirstRow, FirstColumn);
            set
            {
                _sheet.Workbook.Changed();
                string format = _sheet.Workbook.ToText(value);
                XElement sheetData = _sheet.SheetData;
                foreach (var (row, column) in AllCells())
                    Content.SetNumberFormat(sheetData, row, column, format);
            }
        }

        [DispId(146)] public object Font => new ScriptonExcelFont(this);

        [DispId(257)] public int Row => FirstRow;

        [DispId(240)] public int Column => FirstColumn;

        /// <summary>
        /// The number of cells (of rows for Rows, of columns for Columns)
        /// </summary>
        [DispId(118)]
        public int Count => _axis switch
        {
            Axis.Rows => LastRow - FirstRow + 1,
            Axis.Columns => LastColumn - FirstColumn + 1,
            _ => (LastRow - FirstRow + 1) * (LastColumn - FirstColumn + 1),
        };

        [DispId(258)] public object Rows => new ScriptonExcelRange(_sheet, FirstRow, FirstColumn, LastRow, LastColumn, Axis.Rows);

        [DispId(241)] public object Columns => new ScriptonExcelRange(_sheet, FirstRow, FirstColumn, LastRow, LastColumn, Axis.Columns);

        [DispId(246)] public object EntireColumn => new ScriptonExcelRange(_sheet, 1, FirstColumn, 1048576, LastColumn, Axis.Columns);

        [DispId(247)] public object EntireRow => new ScriptonExcelRange(_sheet, FirstRow, 1, LastRow, 16384, Axis.Rows);

        /// <summary>
        /// Cells([RowIndex], [ColumnIndex]) - relative to the range
        /// </summary>
        [DispId(238)]
        public object Cells([In, Optional, DefaultParameterValue(null)] object? RowIndex, [In, Optional, DefaultParameterValue(null)] object? ColumnIndex)
        {
            if (ScriptonExcelApplication.IsMissing(RowIndex))
                return this;
            int row = FirstRow + ScriptonExcelApplication.ToInt(RowIndex) - 1;
            int column = FirstColumn + ((ColumnIndex is string letters) ? ScriptonSpreadsheetDocument.ColumnIndex(letters) : ScriptonExcelApplication.ToInt(ColumnIndex)) - 1;
            return new ScriptonExcelRange(_sheet, row, column, row, column);
        }

        /// <summary>
        /// Address([RowAbsolute], [ColumnAbsolute]) - "$B$2", "$A$1:$C$2"
        /// </summary>
        [DispId(236)]
        public string Address([In, Optional, DefaultParameterValue(null)] object? RowAbsolute, [In, Optional, DefaultParameterValue(null)] object? ColumnAbsolute)
        {
            bool rowAbsolute = ScriptonExcelApplication.IsMissing(RowAbsolute) || Convert.ToBoolean(RowAbsolute, CultureInfo.InvariantCulture);
            bool columnAbsolute = ScriptonExcelApplication.IsMissing(ColumnAbsolute) || Convert.ToBoolean(ColumnAbsolute, CultureInfo.InvariantCulture);
            string Cell(int row, int column) => (columnAbsolute ? "$" : "") + ScriptonSpreadsheetDocument.ColumnName(column) + (rowAbsolute ? "$" : "") + row.ToString(CultureInfo.InvariantCulture);
            string first = Cell(FirstRow, FirstColumn);
            return ((FirstRow == LastRow) && (FirstColumn == LastColumn)) ? first : first + ":" + Cell(LastRow, LastColumn);
        }

        /// <summary>
        /// AutoFit - there is no font measurement, so the column widths stay as they are (True, as with Excel)
        /// </summary>
        [DispId(237)]
        public object AutoFit() => true;

        [DispId(111)]
        public object ClearContents()
        {
            Value = null;
            return true;
        }

        [DispId(235)]
        public object Select()
        {
            _sheet.Activate();
            return true;
        }

        [DispId(304)]
        public object Activate() => Select();

        internal bool GetFontFlag(string flag) => Content.GetFontFlag(_sheet.SheetData, FirstRow, FirstColumn, flag);

        internal void SetFontFlag(string flag, bool value)
        {
            _sheet.Workbook.Changed();
            XElement sheetData = _sheet.SheetData;
            foreach (var (row, column) in AllCells())
                Content.SetFontFlag(sheetData, row, column, flag, value);
        }
    }

    [SourceClassName("Font")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonExcelFont : ScriptonIReflectOnClrType
    {
        private readonly ScriptonExcelRange _range;

        public ScriptonExcelFont(ScriptonExcelRange range) : base(typeof(ScriptonExcelFont))
        {
            _range = range;
        }

        [DispId(96)]
        public object Bold
        {
            get => _range.GetFontFlag("b");
            set => _range.SetFontFlag("b", Convert.ToBoolean(value, CultureInfo.InvariantCulture));
        }

        [DispId(101)]
        public object Italic
        {
            get => _range.GetFontFlag("i");
            set => _range.SetFontFlag("i", Convert.ToBoolean(value, CultureInfo.InvariantCulture));
        }
    }
}
