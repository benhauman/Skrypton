using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// A WordprocessingML document (.docx) in memory - the main document part and the operations of the Word object model on it, with Word's text
    /// conventions: a paragraph ends with "\r", a table cell with "\r\a" (Chr(13) &amp; Chr(7)) and a table row with an extra "\r\a"; a tab is
    /// "\t" and a line break Chr(11).
    /// </summary>
    internal sealed class ScriptonWordprocessingDocument
    {
        internal static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace Xml = XNamespace.Xml;

        private const string CellEnd = "\r\a";

        private readonly ScriptonOpenXmlPackage _package;
        private readonly string _mainPart;
        private readonly XDocument _xml;

        private ScriptonWordprocessingDocument(ScriptonOpenXmlPackage package)
        {
            _package = package;
            _mainPart = package.GetRelationshipTarget("", "/officeDocument")
                ?? throw new System.IO.InvalidDataException("The package has no main document part (it is not a Word document)");
            _xml = package.GetXml(_mainPart);
            Body = _xml.Root?.Element(W + "body") ?? throw new System.IO.InvalidDataException("The main document part has no body (it is not a Word document)");
        }

        public XElement Body { get; }

        public static ScriptonWordprocessingDocument Load(IHostFileSystemHostService fileSystem, string path)
        {
            return new ScriptonWordprocessingDocument(ScriptonOpenXmlPackage.Load(fileSystem, path));
        }

        /// <summary>
        /// A copy of the document (Documents.Add with a template)
        /// </summary>
        public ScriptonWordprocessingDocument Copy()
        {
            ScriptonOpenXmlPackage package = _package.Clone();
            package.SetXml(_mainPart, _xml);
            return new ScriptonWordprocessingDocument(package);
        }

        public void Save(IHostFileSystemHostService fileSystem, string path)
        {
            _package.SetXml(_mainPart, _xml);
            _package.Save(fileSystem, path);
        }

        /// <summary>
        /// A new, empty document (A4, the margins of Word's Normal template)
        /// </summary>
        public static ScriptonWordprocessingDocument CreateEmpty()
        {
            const string contentTypes =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>";
            const string rels =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>";
            const string document =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/>" +
                "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>" +
                "</w:body></w:document>";
            return new ScriptonWordprocessingDocument(ScriptonOpenXmlPackage.Create(new[]
            {
                new KeyValuePair<string, string>("[Content_Types].xml", contentTypes),
                new KeyValuePair<string, string>("_rels/.rels", rels),
                new KeyValuePair<string, string>("word/document.xml", document),
            }));
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Text

        /// <summary>
        /// The text of the whole document (Document.Content.Text)
        /// </summary>
        public string GetContentText() => BlocksText(Body.Elements());

        private static string BlocksText(IEnumerable<XElement> blocks)
        {
            var text = new StringBuilder();
            foreach (XElement block in blocks)
            {
                if (block.Name == W + "p")
                {
                    text.Append(ParagraphText(block)).Append('\r');
                }
                else if (block.Name == W + "tbl")
                {
                    foreach (XElement row in block.Elements(W + "tr"))
                    {
                        foreach (XElement cell in row.Elements(W + "tc"))
                            text.Append(GetCellText(cell));
                        text.Append(CellEnd);
                    }
                }
                else if (block.Name == W + "sdt")
                {
                    text.Append(BlocksText(block.Element(W + "sdtContent")?.Elements() ?? Enumerable.Empty<XElement>()));
                }
            }
            return text.ToString();
        }

        /// <summary>
        /// The text of the cell, with its end of cell mark ("\r\a")
        /// </summary>
        public static string GetCellText(XElement cell)
        {
            string text = BlocksText(cell.Elements());
            if (text.EndsWith("\r", StringComparison.Ordinal))
                text = text.Substring(0, text.Length - 1);
            return text + CellEnd;
        }

        private static string ParagraphText(XElement paragraph)
        {
            var text = new StringBuilder();
            foreach (XElement item in InlineItems(paragraph))
                text.Append(ItemText(item));
            return text.ToString();
        }

        /// <summary>
        /// The text bearing elements and bookmark marks of the paragraph in document order (not the properties, deleted text or text boxes)
        /// </summary>
        private static IEnumerable<XElement> InlineItems(XElement paragraph)
        {
            foreach (XElement element in paragraph.Descendants())
            {
                if (element.Ancestors().TakeWhile(a => a != paragraph).Any(a => (a.Name == W + "pPr") || (a.Name == W + "rPr") || (a.Name == W + "del") || (a.Name == W + "moveFrom") || (a.Name == W + "txbxContent")))
                    continue;
                if ((element.Name == W + "bookmarkStart") || (element.Name == W + "bookmarkEnd"))
                    yield return element;
                else if ((element.Parent?.Name == W + "r") && ((element.Name == W + "t") || (element.Name == W + "tab") || (element.Name == W + "br") || (element.Name == W + "cr") || (element.Name == W + "noBreakHyphen")))
                    yield return element;
            }
        }

        private static string ItemText(XElement item)
        {
            if (item.Name == W + "t")
                return item.Value;
            if (item.Name == W + "tab")
                return "\t";
            if (item.Name == W + "br")
                return ((string?)item.Attribute(W + "type") == "page") ? "\f" : "\v";
            if (item.Name == W + "cr")
                return "\r";
            if (item.Name == W + "noBreakHyphen")
                return "\u001e";
            return "";
        }

        /// <summary>
        /// Word turns vbCrLf, vbCr and vbLf into paragraph marks
        /// </summary>
        private static string[] SplitParagraphs(string text)
        {
            return (text ?? "").Replace("\r\n", "\r").Replace('\n', '\r').Split('\r');
        }

        private static IEnumerable<XElement> CreateRuns(string text, XElement? runProperties)
        {
            var parts = new List<XElement>();
            var current = new StringBuilder();
            void Flush()
            {
                if (current.Length > 0)
                {
                    parts.Add(new XElement(W + "t", new XAttribute(Xml + "space", "preserve"), current.ToString()));
                    current.Clear();
                }
            }
            foreach (char c in text)
            {
                if (c == '\t') { Flush(); parts.Add(new XElement(W + "tab")); }
                else if (c == '\v') { Flush(); parts.Add(new XElement(W + "br")); }
                else if (c == '\f') { Flush(); parts.Add(new XElement(W + "br", new XAttribute(W + "type", "page"))); }
                else if (c == '\a') { /* an end of cell mark is not text */ }
                else current.Append(c);
            }
            Flush();
            if (parts.Count == 0)
                return Enumerable.Empty<XElement>();
            return new[] { new XElement(W + "r", (runProperties == null) ? null : new XElement(runProperties), parts) };
        }

        private static XElement CreateParagraph(XElement? paragraphProperties, XElement? runProperties, string text)
        {
            return new XElement(W + "p", (paragraphProperties == null) ? null : new XElement(paragraphProperties), CreateRuns(text, runProperties));
        }

        private static XElement? FirstRunProperties(XElement? container)
        {
            return container?.Descendants(W + "r").Select(r => r.Element(W + "rPr")).FirstOrDefault(p => p != null);
        }

        /// <summary>
        /// Replace the whole document with the text (Document.Content.Text = ..)
        /// </summary>
        public void SetContentText(string text)
        {
            XElement? template = Body.Elements(W + "p").FirstOrDefault();
            XElement? paragraphProperties = template?.Element(W + "pPr");
            XElement? runProperties = FirstRunProperties(template);
            XElement? sectionProperties = Body.Element(W + "sectPr");
            Body.Elements().Where(e => e != sectionProperties).Remove();
            foreach (string line in SplitParagraphs(TrimFinalParagraphMark(text)))
                AddBeforeSectionProperties(CreateParagraph(paragraphProperties, runProperties, line));
        }

        private void AddBeforeSectionProperties(XElement block)
        {
            XElement? sectionProperties = Body.Element(W + "sectPr");
            if (sectionProperties != null)
                sectionProperties.AddBeforeSelf(block);
            else
                Body.Add(block);
        }

        /// <summary>
        /// The final paragraph mark cannot be replaced (Word keeps it), so "text\r" is "text"
        /// </summary>
        private static string TrimFinalParagraphMark(string text)
        {
            return ((text != null) && text.EndsWith("\r", StringComparison.Ordinal) && !text.EndsWith("\r\a", StringComparison.Ordinal)) ? text.Substring(0, text.Length - 1) : (text ?? "");
        }

        /// <summary>
        /// Replace the content of the cell with the text (Cell.Range.Text = ..) - bookmarks inside the cell are removed (as Word does), except the
        /// bookmarks of whole cells
        /// </summary>
        public static void SetCellText(XElement cell, string text)
        {
            XElement? template = cell.Elements(W + "p").FirstOrDefault();
            XElement? paragraphProperties = template?.Element(W + "pPr");
            XElement? runProperties = FirstRunProperties(template);

            var startIds = new HashSet<string>(cell.Descendants(W + "bookmarkStart").Select(b => (string?)b.Attribute(W + "id") ?? ""));
            List<XElement> keep = cell.Descendants()
                .Where(e => ((e.Name == W + "bookmarkStart") && (e.Attribute(W + "colFirst") != null))
                         || ((e.Name == W + "bookmarkEnd") && !startIds.Contains((string?)e.Attribute(W + "id") ?? "")))
                .ToList();

            cell.Elements().Where(e => e.Name != W + "tcPr").Remove();
            if (text.EndsWith(CellEnd, StringComparison.Ordinal))
                text = text.Substring(0, text.Length - CellEnd.Length);
            foreach (string line in SplitParagraphs(text))
                cell.Add(CreateParagraph(paragraphProperties, runProperties, line));
            XElement first = cell.Element(W + "p")!;
            XElement? firstProperties = first.Element(W + "pPr");
            foreach (XElement mark in keep)
            {
                XElement copy = new XElement(mark);
                if (firstProperties != null)
                    firstProperties.AddAfterSelf(copy);
                else
                    first.AddFirst(copy);
            }
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Bookmarks

        /// <summary>
        /// The names of the bookmarks, sorted by name as Word's Bookmarks collection is (hidden bookmarks - "_GoBack", .. - are not listed)
        /// </summary>
        public List<string> GetBookmarkNames()
        {
            return Body.Descendants(W + "bookmarkStart")
                .Select(b => (string?)b.Attribute(W + "name") ?? "")
                .Where(n => (n.Length > 0) && !n.StartsWith("_", StringComparison.Ordinal))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public XElement? FindBookmarkStart(string name)
        {
            return Body.Descendants(W + "bookmarkStart").FirstOrDefault(b => string.Equals((string?)b.Attribute(W + "name"), name, StringComparison.OrdinalIgnoreCase));
        }

        private XElement? FindBookmarkEnd(XElement start)
        {
            string id = (string?)start.Attribute(W + "id") ?? "";
            return Body.Descendants(W + "bookmarkEnd").FirstOrDefault(b => (string?)b.Attribute(W + "id") == id);
        }

        /// <summary>
        /// The cell that a bookmark of a whole cell marks (a bookmarkStart with w:colFirst in the row of the cell), or null for a bookmark in text
        /// </summary>
        public static XElement? GetBookmarkCell(XElement start)
        {
            string? column = (string?)start.Attribute(W + "colFirst");
            if (column == null)
                return null;
            XElement? row = start.Ancestors(W + "tr").FirstOrDefault();
            return row?.Elements(W + "tc").ElementAtOrDefault(int.Parse(column, CultureInfo.InvariantCulture));
        }

        public string GetBookmarkText(XElement start)
        {
            XElement? cell = GetBookmarkCell(start);
            if (cell != null)
                return GetCellText(cell);

            XElement? end = FindBookmarkEnd(start);
            var text = new StringBuilder();
            bool inside = false;
            foreach (XElement paragraph in Body.Descendants(W + "p"))
            {
                foreach (XElement item in InlineItems(paragraph))
                {
                    if (item == start)
                        inside = true;
                    else if (item == end)
                        return text.ToString();
                    else if (inside)
                        text.Append(ItemText(item));
                }
                if (inside)
                    text.Append('\r');
            }
            return text.ToString();
        }

        /// <summary>
        /// Replace the text that the bookmark marks. As with Word: a bookmark that marked text is removed, a collapsed bookmark (an insertion
        /// point) stays in front of the new text, a bookmark of a whole cell stays.
        /// </summary>
        public void SetBookmarkText(XElement start, string text)
        {
            XElement? cell = GetBookmarkCell(start);
            if (cell != null)
            {
                SetCellText(cell, text);
                return;
            }

            XElement end = FindBookmarkEnd(start) ?? throw new InvalidOperationException($"The bookmark '{(string?)start.Attribute(W + "name")}' has no end");
            if (start.Parent != end.Parent)
                throw new NotSupportedException($"Range.Text: replacing the text of the bookmark '{(string?)start.Attribute(W + "name")}' is only supported if it is inside one paragraph");

            List<XNode> between = start.NodesAfterSelf().TakeWhile(n => n != end).ToList();
            bool collapsed = !between.OfType<XElement>().Any(e => e.DescendantsAndSelf().Any(d => (d.Name == W + "t") || (d.Name == W + "tab") || (d.Name == W + "br")));
            XElement? runProperties = between.OfType<XElement>().Where(e => e.Name == W + "r").Select(r => r.Element(W + "rPr")).FirstOrDefault(p => p != null)
                ?? start.ElementsBeforeSelf(W + "r").Select(r => r.Element(W + "rPr")).LastOrDefault(p => p != null)
                ?? end.ElementsAfterSelf(W + "r").Select(r => r.Element(W + "rPr")).FirstOrDefault(p => p != null);
            foreach (XNode node in between)
                node.Remove();

            // the new text goes after the bookmark (a collapsed bookmark stays in front of it) - or where it was (the bookmark is removed)
            XElement insertAfter = end;
            string[] lines = SplitParagraphs(text);
            List<XElement> runs = CreateRuns(lines[0], runProperties).ToList();
            foreach (XElement run in runs)
            {
                insertAfter.AddAfterSelf(run);
                insertAfter = run;
            }
            if (lines.Length > 1)
                SplitParagraphAfter(insertAfter, lines.Skip(1).ToList(), runProperties);

            if (!collapsed)
            {
                start.Remove();
                end.Remove();
            }
        }

        /// <summary>
        /// The text after the element moves to new paragraphs (the lines), the rest of the paragraph after the element goes to the last of them
        /// </summary>
        private static void SplitParagraphAfter(XElement element, List<string> lines, XElement? runProperties)
        {
            XElement paragraph = element.Ancestors(W + "p").First();
            if (element.Parent != paragraph)
                throw new NotSupportedException("Range.Text: inserting paragraphs is only supported in plain paragraphs (not in hyperlinks or fields)");
            XElement? paragraphProperties = paragraph.Element(W + "pPr");
            List<XNode> rest = element.NodesAfterSelf().ToList();
            foreach (XNode node in rest)
                node.Remove();
            XElement previous = paragraph;
            XElement? last = null;
            foreach (string line in lines)
            {
                last = CreateParagraph(paragraphProperties, runProperties, line);
                previous.AddAfterSelf(last);
                previous = last;
            }
            last!.Add(rest);
        }

        /// <summary>
        /// Add a bookmark (Bookmarks.Add) - an existing bookmark with the name is replaced
        /// </summary>
        public void AddBookmark(string name, ScriptonWordRangeTarget target)
        {
            if (FindBookmarkStart(name) is XElement existing)
                RemoveBookmark(existing);
            string id = (Body.Descendants(W + "bookmarkStart").Select(b => int.TryParse((string?)b.Attribute(W + "id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int i) ? i : -1).DefaultIfEmpty(-1).Max() + 1)
                .ToString(CultureInfo.InvariantCulture);
            var start = new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name));
            var end = new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id));

            XElement? cell = target.Cell ?? ((target.BookmarkStart != null) ? GetBookmarkCell(target.BookmarkStart) : null);
            if (cell != null)
            {
                // a bookmark of a whole cell: its start is in the first cell of the row (with the column), its end after the row (as Word writes it)
                XElement row = cell.Parent!;
                int column = row.Elements(W + "tc").ToList().IndexOf(cell);
                start.Add(new XAttribute(W + "colFirst", column), new XAttribute(W + "colLast", column));
                XElement firstParagraph = row.Elements(W + "tc").First().Element(W + "p") ?? throw new InvalidOperationException("The cell has no paragraph");
                InsertAtParagraphStart(firstParagraph, start);
                row.AddAfterSelf(end);
            }
            else if (target.BookmarkStart != null)
            {
                target.BookmarkStart.AddBeforeSelf(start);
                (FindBookmarkEnd(target.BookmarkStart) ?? target.BookmarkStart).AddAfterSelf(end);
            }
            else
            {
                // the whole document
                List<XElement> paragraphs = Body.Elements(W + "p").ToList();
                if (paragraphs.Count == 0)
                {
                    AddBeforeSectionProperties(new XElement(W + "p"));
                    paragraphs = Body.Elements(W + "p").ToList();
                }
                InsertAtParagraphStart(paragraphs[0], start);
                paragraphs[paragraphs.Count - 1].Add(end);
            }
        }

        private static void InsertAtParagraphStart(XElement paragraph, XElement element)
        {
            XElement? properties = paragraph.Element(W + "pPr");
            if (properties != null)
                properties.AddAfterSelf(element);
            else
                paragraph.AddFirst(element);
        }

        public void RemoveBookmark(XElement start)
        {
            FindBookmarkEnd(start)?.Remove();
            start.Remove();
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Tables

        /// <summary>
        /// The tables of the document (not the tables nested in tables, as Word's Document.Tables)
        /// </summary>
        public List<XElement> GetTables() => Body.Descendants(W + "tbl").Where(t => !t.Ancestors(W + "tbl").Any()).ToList();

        /// <summary>
        /// Replace the whole document with an empty table (Tables.Add(Document.Content, ..)) - Word keeps a paragraph after the table
        /// </summary>
        public XElement ReplaceContentWithTable(int numRows, int numColumns)
        {
            const int textWidth = 9026; // A4 width without the margins (twips)
            XElement? sectionProperties = Body.Element(W + "sectPr");
            XElement? paragraphProperties = Body.Elements(W + "p").FirstOrDefault()?.Element(W + "pPr");
            Body.Elements().Where(e => e != sectionProperties).Remove();

            int columnWidth = textWidth / numColumns;
            var table = new XElement(W + "tbl",
                new XElement(W + "tblPr",
                    new XElement(W + "tblW", new XAttribute(W + "w", 0), new XAttribute(W + "type", "auto")),
                    new XElement(W + "tblLook", new XAttribute(W + "val", "0000"))),
                new XElement(W + "tblGrid", Enumerable.Range(0, numColumns).Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", columnWidth)))),
                Enumerable.Range(0, numRows).Select(_ => new XElement(W + "tr",
                    Enumerable.Range(0, numColumns).Select(__ => new XElement(W + "tc",
                        new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", columnWidth), new XAttribute(W + "type", "dxa"))),
                        new XElement(W + "p", (paragraphProperties == null) ? null : new XElement(paragraphProperties)))))));
            AddBeforeSectionProperties(table);
            AddBeforeSectionProperties(new XElement(W + "p", (paragraphProperties == null) ? null : new XElement(paragraphProperties)));
            return table;
        }

        /// <summary>
        /// Add a row (Rows.Add) like the row it is added before (or the last row): the cells keep their formatting, not their text or bookmarks
        /// </summary>
        public static XElement AddRow(XElement table, XElement? beforeRow)
        {
            XElement template = beforeRow ?? table.Elements(W + "tr").Last();
            var row = new XElement(template);
            row.Descendants().Where(e => (e.Name == W + "bookmarkStart") || (e.Name == W + "bookmarkEnd")).Remove();
            foreach (XElement cell in row.Elements(W + "tc"))
            {
                XElement? first = cell.Elements(W + "p").FirstOrDefault();
                XElement? paragraphProperties = first?.Element(W + "pPr");
                cell.Elements().Where(e => e.Name != W + "tcPr").Remove();
                cell.Add(new XElement(W + "p", (paragraphProperties == null) ? null : new XElement(paragraphProperties)));
            }
            if (beforeRow != null)
                beforeRow.AddBeforeSelf(row);
            else
                table.Elements(W + "tr").Last().AddAfterSelf(row);
            return row;
        }
    }

    /// <summary>
    /// What a Range covers: the whole document, a cell or a bookmark
    /// </summary>
    internal sealed class ScriptonWordRangeTarget
    {
        private ScriptonWordRangeTarget(XElement? cell, XElement? bookmarkStart)
        {
            Cell = cell;
            BookmarkStart = bookmarkStart;
        }

        public static readonly ScriptonWordRangeTarget Content = new ScriptonWordRangeTarget(null, null);

        public static ScriptonWordRangeTarget ForCell(XElement cell) => new ScriptonWordRangeTarget(cell, null);

        public static ScriptonWordRangeTarget ForBookmark(XElement bookmarkStart) => new ScriptonWordRangeTarget(null, bookmarkStart);

        public XElement? Cell { get; }
        public XElement? BookmarkStart { get; }
        public bool IsContent => (Cell == null) && (BookmarkStart == null);
    }
}
