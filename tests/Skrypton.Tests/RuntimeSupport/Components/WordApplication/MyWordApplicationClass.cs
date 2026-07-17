using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.WordApplication
{
    [DefaultMember("Name")]
    internal sealed class MyWordApplicationClass : IReflectOnClrType
    {
        /* Add COM-Reference "Microsoft Word 16.0 Object Library" (Microsoft.Office.Interop.Word) <Guid>00020905-0000-0000-c000-000000000046</Guid>
         * => "projects.ToDelete\ConsoleApp2\bin\Debug\net10.0\Interop.Microsoft.Office.Interop.Word.dll"

         */

        private readonly IRuntimeHost _runtimeHost;

        public MyWordApplicationClass(IRuntimeHost runtimeHost)
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        [DispId(0)][IsDefault] public string Name { get; set; }

        [DispId(80)] public string Caption { get; set; }

        [DispId(23)] public bool Visible { get; set; }


        [DispId(6)]
        public Documents Documents => new MyWordDocuments(_documents);

        private readonly List<MyWordDocument> _documents = new List<MyWordDocument>();
    }

    internal interface Documents : System.Collections.IEnumerable
    {
    }

    internal interface Document// : _Document, DocumentEvents2_Event
    {
    }
    internal sealed class MyWordDocuments : IReflectOnClrType, Documents
    {
        private readonly List<MyWordDocument> _documents;

        public MyWordDocuments(List<MyWordDocument> documents)
        {
            _documents = documents ?? throw new ArgumentNullException(nameof(documents));
        }

        public IEnumerator GetEnumerator()
        {
            return _documents.GetEnumerator();
        }

        [DispId(19)]
        public Document Open([In] object FileName) => OpenCore(FileName);
        private Document OpenCore([In] object FileName)//, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ConfirmConversions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnly, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PasswordDocument, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PasswordTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Revert, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePasswordDocument, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePasswordTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Format, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Visible, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OpenAndRepair, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DocumentDirection, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object NoEncodingDialog, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object XMLTransform)
        {
            Console.WriteLine($"[WORD].Open('{FileName}')");
            // VBScript passes Windows-style paths (backslashes). Path.GetFileName only
            // treats '\' as a separator on Windows, so split on both separators to get
            // the file name consistently on Linux and Windows.
            string name = ((string)FileName).Split('/', '\\').Last();
            MyWordDocument doc = new MyWordDocument(name);


            if (name == "A-XXX-VL-Firma-hL_Cons.docx") // CT98_dialog287_ButtonCreateWord_Click
            {
                var t1 = new MyWordTableClass();
                var t1r1 = new MyWordTableRowClass();
                t1._rows.Add(t1r1);
                var t1r2 = new MyWordTableRowClass();
                t1r2._cells.Add(new MyWordTableRowCellClass());
                t1r2._cells.Add(new MyWordTableRowCellClass());
                t1r2._cells.Add(new MyWordTableRowCellClass());
                t1r2._cells.Add(new MyWordTableRowCellClass());
                t1r2._cells.Add(new MyWordTableRowCellClass());
                t1r2._cells.Add(new MyWordTableRowCellClass());

                t1._rows.Add(t1r2);
                doc._tables.Add(t1);
            }

            _documents.Add(doc);
            return doc;
        }
    }

    internal sealed class MyWordDocument : IReflectOnClrType, Document
    {
        internal readonly List<MyWordTableClass> _tables = new List<MyWordTableClass>();
        internal readonly List<MyWordBookmark> _bookmarks = new List<MyWordBookmark>();
        public MyWordDocument(string name)
        {
            Name = name;
        }

        [DispId(0)] public string Name { get; }


        [DispId(6)] public Tables Tables() => new MyWordTables(_tables);
        [DispId(6)] public Table Tables(object index) => new MyWordTables(_tables).GetTableByIndex(index);


        [DispId(4)] public Bookmarks Bookmarks() => new MyWordBookmarks(_bookmarks);
        [DispId(4)] public Bookmark Bookmarks(object name) => new MyWordBookmarks(_bookmarks).GetBookmarkByName(name);
    }



    internal sealed class MyWordTables : IReflectOnClrType, Tables
    {
        private readonly List<MyWordTableClass> _tables;

        public MyWordTables(List<MyWordTableClass> tables)
        {
            _tables = tables ?? throw new ArgumentNullException(nameof(tables));
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        [DispId(0)]
        public MyWordTableClass GetTableByIndex(object index) // Word collections are 1-based, just like VBScript.
        {
            int idx = Convert.ToInt32(index);
            if (idx <= 0)
                throw new ArgumentException("Index cannot be zero. Word collections are 1-based", nameof(index));
            return _tables[idx - 1];
        }
    }

    internal interface Bookmarks : IEnumerable
    {
    }

    [DefaultMember("Name")]
    internal interface Bookmark // A bookmark is a named marker inside a Word document that points to a specific location or range of text.
    {

    }
    internal interface Tables : IEnumerable
    {
    }

    internal interface Table
    {
    }
    internal interface Range
    {
        // a Range represents a contiguous area of a document.
    }
    internal interface Rows : IEnumerable
    {

    }
    internal interface Row
    {

    }
    internal interface Cells : IEnumerable
    {

    }
    internal interface Cell
    {

    }
    internal sealed class MyWordTableClass : IReflectOnClrType, Table
    {
        internal MyWordRange _range = new MyWordRange(2);
        internal readonly List<MyWordTableRowClass> _rows = new List<MyWordTableRowClass>();

        public MyWordTableClass()
        {

        }

        [DispId(0)] public Skrypton.Tests.RuntimeSupport.Components.WordApplication.Range Range => _range;
        [DispId(101)] public Rows Rows() => new MyWordTableRowsClass(_rows);
        [DispId(101)] public Row Rows(object index) => new MyWordTableRowsClass(_rows).GetRowByIndex(index);
    }

    [DefaultMember("Text")]
    internal sealed class MyWordRange : IReflectOnClrType, Skrypton.Tests.RuntimeSupport.Components.WordApplication.Range
    {
        private readonly int _id;

        public MyWordRange(int id)
        {
            _id = id;
        }

        private string _text = "";
        [DispId(0)]
        [IsDefault]
        public object Text
        {
            get => _text;
            set => _text = value?.ToString();
        }
    }

    internal sealed class MyWordTableRowsClass : IReflectOnClrType, Rows
    {
        private readonly List<MyWordTableRowClass> _rows;

        public MyWordTableRowsClass(List<MyWordTableRowClass> _rows)
        {
            this._rows = _rows;
        }

        public IEnumerator GetEnumerator()
        {
            return _rows.GetEnumerator();
        }

        [DispId(0)]
        public Row GetRowByIndex(object index)
        {
            int idx = Convert.ToInt32(index);
            if (idx <= 0)
                throw new ArgumentException("Index cannot be zero. Word collections are 1-based", nameof(index));
            return _rows[idx - 1];
        }
    }

    internal sealed class MyWordTableRowClass : IReflectOnClrType, Row
    {
        internal readonly List<MyWordTableRowCellClass> _cells = new List<MyWordTableRowCellClass>();
        public MyWordTableRowClass()
        {

        }

        [DispId(100)] public Cells Cells() => new MyWordTableRowCellsClass(_cells);
        [DispId(100)] public Cell Cells(int index) => new MyWordTableRowCellsClass(_cells).GetCellByIndex(index);
    }
    internal sealed class MyWordTableRowCellsClass : IReflectOnClrType, Cells
    {
        private readonly List<MyWordTableRowCellClass> _cells;

        public MyWordTableRowCellsClass(List<MyWordTableRowCellClass> cells)
        {
            _cells = cells;
        }

        public IEnumerator GetEnumerator()
        {
            return _cells.GetEnumerator();
        }
        [DispId(0)]
        public Cell GetCellByIndex(object index)
        {
            int idx = Convert.ToInt32(index);
            if (idx <= 0)
                throw new ArgumentException("Index cannot be zero. Word collections are 1-based", nameof(index));
            return _cells[idx - 1];
        }

    }

    [DefaultMember("Range")]
    internal sealed class MyWordTableRowCellClass : IReflectOnClrType, Cell
    {
        private MyWordRange _range = new MyWordRange(1);
        public MyWordTableRowCellClass()
        {

        }

        [DispId(0)]
        [IsDefault]
        public object Range
        {
            get => _range.Text;
            set => _range.Text = value;
        }
    }

    internal sealed class MyWordBookmarks : IReflectOnClrType, Bookmarks
    {
        private readonly List<MyWordBookmark> _bookmarks;

        public MyWordBookmarks(List<MyWordBookmark> bookmarks)
        {
            _bookmarks = bookmarks;
        }
        public IEnumerator GetEnumerator()
        {
            return _bookmarks.GetEnumerator();
        }

        public Bookmark GetBookmarkByName(object name)
        {
            string bookmarkName = (string)name;
            var bookmark = _bookmarks.SingleOrDefault(x => string.Equals(x.Name, bookmarkName, StringComparison.OrdinalIgnoreCase));
            if (bookmark == null)
            {
                bookmark = new MyWordBookmark(bookmarkName);
                //bookmarks.Add(bookmark);
            }
            return bookmark;
        }
    }

    [DefaultMember("Name")]
    internal sealed class MyWordBookmark : IReflectOnClrType, Bookmark
    {
        private MyWordRange _range = new MyWordRange(3);
        public MyWordBookmark(string name)
        {
            Name = name;
        }

        [DispId(0)][IsDefault] public string Name { get; }
        [DispId(1)] public Skrypton.Tests.RuntimeSupport.Components.WordApplication.Range Range => _range;
    }

}
