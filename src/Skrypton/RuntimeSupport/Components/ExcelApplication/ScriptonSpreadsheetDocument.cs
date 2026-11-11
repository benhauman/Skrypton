using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// A SpreadsheetML workbook (.xlsx) in memory. The sheets are changed in place (the formatting, formulas and everything else of a template are
    /// kept): values are written as numbers, inline strings, booleans or dates (numbers with a date format), a formula ("=..") as a formula that
    /// Excel calculates when the workbook is opened.
    /// </summary>
    internal sealed class ScriptonSpreadsheetDocument
    {
        internal static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

        private readonly ScriptonOpenXmlPackage _package;
        private readonly string _workbookPart;
        private readonly XDocument _workbook;
        private readonly XDocument _workbookRelationships;
        private readonly XDocument? _styles;
        private readonly List<string> _sharedStrings;
        private readonly Dictionary<string, XDocument> _sheetXml = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);

        private ScriptonSpreadsheetDocument(ScriptonOpenXmlPackage package)
        {
            _package = package;
            _workbookPart = package.GetRelationshipTarget("", "/officeDocument") ?? throw new System.IO.InvalidDataException("The package has no workbook part (it is not an Excel workbook)");
            _workbook = package.GetXml(_workbookPart);
            if (_workbook.Root?.Name != S + "workbook")
                throw new System.IO.InvalidDataException("The main part is not a workbook (it is not an Excel workbook)");
            _workbookRelationships = package.GetXml(RelationshipsPart(_workbookPart));
            string? stylesPart = package.GetRelationshipTarget(_workbookPart, "/styles");
            _styles = (stylesPart == null) ? null : package.GetXml(stylesPart);
            string? sharedStringsPart = package.GetRelationshipTarget(_workbookPart, "/sharedStrings");
            _sharedStrings = (sharedStringsPart == null)
                ? new List<string>()
                : package.GetXml(sharedStringsPart).Root!.Elements(S + "si").Select(StringItemText).ToList();
            ActiveSheetIndex = (int?)_workbook.Root.Element(S + "bookViews")?.Element(S + "workbookView")?.Attribute("activeTab") ?? 0;
        }

        public static ScriptonSpreadsheetDocument Load(IHostFileSystemHostService fileSystem, string path) => new ScriptonSpreadsheetDocument(ScriptonOpenXmlPackage.Load(fileSystem, path));

        public ScriptonSpreadsheetDocument Copy()
        {
            WriteParts();
            return new ScriptonSpreadsheetDocument(_package.Clone());
        }

        public void Save(IHostFileSystemHostService fileSystem, string path)
        {
            WriteParts();
            _package.Save(fileSystem, path);
        }

        private void WriteParts()
        {
            // Excel calculates the formulas (without cached values) when the workbook is opened
            XElement root = _workbook.Root!;
            XElement? calculation = root.Element(S + "calcPr");
            if (calculation == null)
            {
                calculation = new XElement(S + "calcPr");
                XElement? after = root.Elements().LastOrDefault(e => (e.Name == S + "sheets") || (e.Name == S + "functionGroups") || (e.Name == S + "externalReferences") || (e.Name == S + "definedNames"));
                if (after != null)
                    after.AddAfterSelf(calculation);
                else
                    root.Add(calculation);
            }
            calculation.SetAttributeValue("fullCalcOnLoad", "1");
            XElement? view = root.Element(S + "bookViews")?.Element(S + "workbookView");
            view?.SetAttributeValue("activeTab", ActiveSheetIndex);

            _package.SetXml(_workbookPart, _workbook);
            _package.SetXml(RelationshipsPart(_workbookPart), _workbookRelationships);
            string? stylesPart = _package.GetRelationshipTarget(_workbookPart, "/styles");
            if ((stylesPart != null) && (_styles != null))
                _package.SetXml(stylesPart, _styles);
            foreach (var sheet in _sheetXml)
                _package.SetXml(sheet.Key, sheet.Value);
        }

        /// <summary>
        /// A new workbook with one sheet ("Sheet1", as Excel)
        /// </summary>
        public static ScriptonSpreadsheetDocument CreateEmpty()
        {
            const string declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
            const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return new ScriptonSpreadsheetDocument(ScriptonOpenXmlPackage.Create(new[]
            {
                new KeyValuePair<string, string>("[Content_Types].xml", declaration +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"" + WorksheetContentType + "\"/>" +
                    "</Types>"),
                new KeyValuePair<string, string>("_rels/.rels", declaration +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>"),
                new KeyValuePair<string, string>("xl/workbook.xml", declaration +
                    "<workbook xmlns=\"" + main + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<bookViews><workbookView activeTab=\"0\"/></bookViews>" +
                    "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                    "</workbook>"),
                new KeyValuePair<string, string>("xl/_rels/workbook.xml.rels", declaration +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"" + WorksheetRelationshipType + "\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>"),
                new KeyValuePair<string, string>("xl/styles.xml", declaration +
                    "<styleSheet xmlns=\"" + main + "\">" +
                    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>" +
                    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
                    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                    "</styleSheet>"),
                new KeyValuePair<string, string>("xl/worksheets/sheet1.xml", EmptySheetXml()),
            }));
        }

        private static string EmptySheetXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/></worksheet>";
        }

        private static string RelationshipsPart(string part)
        {
            int slash = part.LastIndexOf('/');
            return part.Substring(0, slash + 1) + "_rels/" + part.Substring(slash + 1) + ".rels";
        }

        private static string StringItemText(XElement item)
        {
            // <si><t>..</t></si> or rich text <si><r><t>..</t></r>..</si> (phonetic runs are not text)
            return string.Concat(item.Descendants(S + "t").Where(t => !t.Ancestors(S + "rPh").Any()).Select(t => t.Value));
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Sheets

        public int ActiveSheetIndex { get; set; }

        private XElement SheetsElement => _workbook.Root!.Element(S + "sheets") ?? throw new System.IO.InvalidDataException("The workbook has no sheets");

        /// <summary>
        /// The worksheets in their order (chart sheets are not worksheets)
        /// </summary>
        public List<XElement> GetSheetElements() => SheetsElement.Elements(S + "sheet").ToList();

        public static string GetSheetName(XElement sheet) => (string?)sheet.Attribute("name") ?? "";

        public static void SetSheetName(XElement sheet, string name) => sheet.SetAttributeValue("name", name);

        public XElement GetSheetData(XElement sheet)
        {
            string id = (string?)sheet.Attribute(R + "id") ?? "";
            string target = (string?)_workbookRelationships.Root!.Elements(Rel + "Relationship").FirstOrDefault(r => (string?)r.Attribute("Id") == id)?.Attribute("Target")
                ?? throw new System.IO.InvalidDataException($"The sheet '{GetSheetName(sheet)}' has no part");
            string part = target.StartsWith("/", StringComparison.Ordinal) ? target.Substring(1) : _workbookPart.Substring(0, _workbookPart.LastIndexOf('/') + 1) + target;
            if (!_sheetXml.TryGetValue(part, out XDocument? xml))
            {
                xml = _package.GetXml(part);
                _sheetXml[part] = xml;
            }
            XElement worksheet = xml.Root!;
            XElement? sheetData = worksheet.Element(S + "sheetData");
            if (sheetData == null)
            {
                sheetData = new XElement(S + "sheetData");
                worksheet.Add(sheetData);
            }
            return sheetData;
        }

        /// <summary>
        /// Add a worksheet before the sheet at the position (Worksheets.Add) - it is named "SheetN" with the next free number
        /// </summary>
        public XElement AddSheet(int beforeIndex)
        {
            int number = 1;
            var names = new HashSet<string>(GetSheetElements().Select(GetSheetName), StringComparer.OrdinalIgnoreCase);
            while (names.Contains("Sheet" + number.ToString(CultureInfo.InvariantCulture)))
                number++;

            int partNumber = 1;
            while (_package.HasPart("xl/worksheets/sheet" + partNumber.ToString(CultureInfo.InvariantCulture) + ".xml") || _sheetXml.ContainsKey("xl/worksheets/sheet" + partNumber.ToString(CultureInfo.InvariantCulture) + ".xml"))
                partNumber++;
            string part = "xl/worksheets/sheet" + partNumber.ToString(CultureInfo.InvariantCulture) + ".xml";
            _sheetXml[part] = XDocument.Parse(EmptySheetXml());

            XElement relationships = _workbookRelationships.Root!;
            int rid = 1;
            while (relationships.Elements(Rel + "Relationship").Any(r => (string?)r.Attribute("Id") == "rId" + rid.ToString(CultureInfo.InvariantCulture)))
                rid++;
            string relationshipId = "rId" + rid.ToString(CultureInfo.InvariantCulture);
            relationships.Add(new XElement(Rel + "Relationship", new XAttribute("Id", relationshipId), new XAttribute("Type", WorksheetRelationshipType),
                new XAttribute("Target", (_workbookPart.StartsWith("xl/", StringComparison.Ordinal) ? "worksheets/sheet" : "/xl/worksheets/sheet") + partNumber.ToString(CultureInfo.InvariantCulture) + ".xml")));

            XDocument contentTypes = _package.GetXml("[Content_Types].xml");
            contentTypes.Root!.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", WorksheetContentType)));
            _package.SetXml("[Content_Types].xml", contentTypes);

            int sheetId = GetSheetElements().Select(s => (int?)s.Attribute("sheetId") ?? 0).DefaultIfEmpty(0).Max() + 1;
            var sheet = new XElement(S + "sheet", new XAttribute("name", "Sheet" + number.ToString(CultureInfo.InvariantCulture)), new XAttribute("sheetId", sheetId), new XAttribute(R + "id", relationshipId));
            List<XElement> sheets = GetSheetElements();
            if (beforeIndex < sheets.Count)
                sheets[beforeIndex].AddBeforeSelf(sheet);
            else
                SheetsElement.Add(sheet);
            return sheet;
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Cells

        /// <summary>
        /// The value of the cell: a Double, a String, a Boolean, a DateTime (a number with a date format), an error text ("#DIV/0!") or null (empty)
        /// </summary>
        public object? GetValue(XElement sheetData, int row, int column)
        {
            XElement? cell = FindCell(sheetData, row, column, create: false);
            if (cell == null)
                return null;
            string? type = (string?)cell.Attribute("t");
            string? value = (string?)cell.Element(S + "v");
            switch (type)
            {
                case "inlineStr":
                    XElement? inline = cell.Element(S + "is");
                    return (inline == null) ? "" : StringItemText(inline);
                case "s":
                    return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) && (index >= 0) && (index < _sharedStrings.Count) ? _sharedStrings[index] : "";
                case "str":
                case "e":
                    return value ?? "";
                case "b":
                    return value == "1";
            }
            if (value == null)
                return null;
            double number = double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
            return IsDateFormat(GetNumberFormat(cell)) ? (object)DateTime.FromOADate(number) : number;
        }

        public static string? GetFormula(XElement sheetData, int row, int column) => (string?)FindCell(sheetData, row, column, create: false)?.Element(S + "f");

        /// <summary>
        /// Set the value (null clears the cell; a string starting with "=" is a formula)
        /// </summary>
        public void SetValue(XElement sheetData, int row, int column, object? value)
        {
            XElement cell = FindCell(sheetData, row, column, create: true)!;
            cell.Attribute("t")?.Remove();
            cell.Elements().Remove();
            switch (value)
            {
                case null:
                    break;
                case string formula when formula.StartsWith("=", StringComparison.Ordinal) && (formula.Length > 1):
                    cell.Add(new XElement(S + "f", formula.Substring(1)));
                    break;
                case string text:
                    cell.SetAttributeValue("t", "inlineStr");
                    cell.Add(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
                    break;
                case bool boolean:
                    cell.SetAttributeValue("t", "b");
                    cell.Add(new XElement(S + "v", boolean ? "1" : "0"));
                    break;
                case DateTime date:
                    if (!IsDateFormat(GetNumberFormat(cell)))
                        SetNumberFormat(cell, (date.TimeOfDay == TimeSpan.Zero) ? "m/d/yyyy" : "m/d/yyyy h:mm");
                    cell.Add(new XElement(S + "v", date.ToOADate().ToString("R", CultureInfo.InvariantCulture)));
                    break;
                case double number:
                    cell.Add(new XElement(S + "v", number.ToString("R", CultureInfo.InvariantCulture)));
                    break;
                default:
                    throw new ArgumentException($"Unsupported cell value of type {value.GetType().Name}", nameof(value));
            }
        }

        /// <summary>
        /// The used cells of the sheet: first row, first column, last row, last column (all 1 if the sheet is empty)
        /// </summary>
        public static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) GetUsedRange(XElement sheetData)
        {
            List<(int Row, int Column)> used = sheetData.Elements(S + "row").Elements(S + "c")
                .Where(c => c.Elements().Any())
                .Select(c => ParseAddress((string?)c.Attribute("r") ?? "A1"))
                .ToList();
            if (used.Count == 0)
                return (1, 1, 1, 1);
            return (used.Min(c => c.Row), used.Min(c => c.Column), used.Max(c => c.Row), used.Max(c => c.Column));
        }

        private static XElement? FindCell(XElement sheetData, int row, int column, bool create)
        {
            string rowText = row.ToString(CultureInfo.InvariantCulture);
            XElement? rowElement = sheetData.Elements(S + "row").FirstOrDefault(r => (string?)r.Attribute("r") == rowText);
            if (rowElement == null)
            {
                if (!create)
                    return null;
                rowElement = new XElement(S + "row", new XAttribute("r", rowText));
                XElement? next = sheetData.Elements(S + "row").FirstOrDefault(r => ((int?)r.Attribute("r") ?? 0) > row);
                if (next != null)
                    next.AddBeforeSelf(rowElement);
                else
                    sheetData.Add(rowElement);
            }
            string reference = ColumnName(column) + rowText;
            XElement? cell = rowElement.Elements(S + "c").FirstOrDefault(c => string.Equals((string?)c.Attribute("r"), reference, StringComparison.OrdinalIgnoreCase));
            if ((cell == null) && create)
            {
                cell = new XElement(S + "c", new XAttribute("r", reference));
                XElement? next = rowElement.Elements(S + "c").FirstOrDefault(c => ParseAddress((string?)c.Attribute("r") ?? "A1").Column > column);
                if (next != null)
                    next.AddBeforeSelf(cell);
                else
                    rowElement.Add(cell);
                rowElement.Attribute("spans")?.Remove(); // optional, Excel recalculates it
            }
            return cell;
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Styles (number formats, bold, italic)

        private static readonly Dictionary<int, string> BuiltInNumberFormats = new Dictionary<int, string>
        {
            { 0, "General" }, { 1, "0" }, { 2, "0.00" }, { 3, "#,##0" }, { 4, "#,##0.00" }, { 9, "0%" }, { 10, "0.00%" }, { 11, "0.00E+00" },
            { 12, "# ?/?" }, { 13, "# ??/??" }, { 14, "m/d/yyyy" }, { 15, "d-mmm-yy" }, { 16, "d-mmm" }, { 17, "mmm-yy" }, { 18, "h:mm AM/PM" },
            { 19, "h:mm:ss AM/PM" }, { 20, "h:mm" }, { 21, "h:mm:ss" }, { 22, "m/d/yyyy h:mm" }, { 37, "#,##0 ;(#,##0)" }, { 38, "#,##0 ;[Red](#,##0)" },
            { 39, "#,##0.00;(#,##0.00)" }, { 40, "#,##0.00;[Red](#,##0.00)" }, { 45, "mm:ss" }, { 46, "[h]:mm:ss" }, { 47, "mmss.0" }, { 48, "##0.0E+0" },
            { 49, "@" },
        };

        private XElement StyleSheet => _styles?.Root ?? throw new NotSupportedException("The workbook has no styles part, so formats cannot be changed");

        private XElement CellFormats => StyleSheet.Element(S + "cellXfs") ?? throw new System.IO.InvalidDataException("The styles have no cell formats");

        private XElement? CellFormat(XElement cell)
        {
            int index = (int?)cell.Attribute("s") ?? 0;
            return _styles?.Root?.Element(S + "cellXfs")?.Elements(S + "xf").ElementAtOrDefault(index);
        }

        public string GetNumberFormat(XElement? cell)
        {
            if (cell == null)
                return "General";
            int id = (int?)CellFormat(cell)?.Attribute("numFmtId") ?? 0;
            if (BuiltInNumberFormats.TryGetValue(id, out string? format))
                return format;
            return (string?)_styles?.Root?.Element(S + "numFmts")?.Elements(S + "numFmt").FirstOrDefault(n => (int?)n.Attribute("numFmtId") == id)?.Attribute("formatCode") ?? "General";
        }

        public string GetNumberFormat(XElement sheetData, int row, int column) => GetNumberFormat(FindCell(sheetData, row, column, create: false));

        public void SetNumberFormat(XElement sheetData, int row, int column, string format) => SetNumberFormat(FindCell(sheetData, row, column, create: true)!, format);

        private void SetNumberFormat(XElement cell, string format)
        {
            int id = BuiltInNumberFormats.Where(f => f.Value == format).Select(f => (int?)f.Key).FirstOrDefault() ?? GetOrAddCustomNumberFormat(format);
            ChangeCellFormat(cell, xf => xf.SetAttributeValue("numFmtId", id), xf => xf.SetAttributeValue("applyNumberFormat", "1"));
        }

        private int GetOrAddCustomNumberFormat(string format)
        {
            XElement? numberFormats = StyleSheet.Element(S + "numFmts");
            if (numberFormats == null)
            {
                numberFormats = new XElement(S + "numFmts", new XAttribute("count", 0));
                StyleSheet.AddFirst(numberFormats);
            }
            XElement? existing = numberFormats.Elements(S + "numFmt").FirstOrDefault(n => (string?)n.Attribute("formatCode") == format);
            if (existing != null)
                return (int)existing.Attribute("numFmtId")!;
            int id = Math.Max(164, numberFormats.Elements(S + "numFmt").Select(n => (int?)n.Attribute("numFmtId") ?? 0).DefaultIfEmpty(0).Max() + 1);
            numberFormats.Add(new XElement(S + "numFmt", new XAttribute("numFmtId", id), new XAttribute("formatCode", format)));
            numberFormats.SetAttributeValue("count", numberFormats.Elements(S + "numFmt").Count());
            return id;
        }

        public bool GetFontFlag(XElement sheetData, int row, int column, string flag)
        {
            XElement? cell = FindCell(sheetData, row, column, create: false);
            XElement? font = (cell == null) ? DefaultFont() : Font((int?)CellFormat(cell)?.Attribute("fontId") ?? 0);
            XElement? element = font?.Element(S + flag);
            return (element != null) && ((string?)element.Attribute("val") ?? "1") != "0";
        }

        public void SetFontFlag(XElement sheetData, int row, int column, string flag, bool value)
        {
            XElement cell = FindCell(sheetData, row, column, create: true)!;
            XElement fonts = StyleSheet.Element(S + "fonts") ?? throw new System.IO.InvalidDataException("The styles have no fonts");
            var font = new XElement(Font((int?)CellFormat(cell)?.Attribute("fontId") ?? 0) ?? new XElement(S + "font"));
            font.Elements(S + flag).Remove();
            if (value)
                font.AddFirst(new XElement(S + flag)); // b / i come first in a font
            int fontId = fonts.Elements(S + "font").ToList().FindIndex(f => XNode.DeepEquals(f, font));
            if (fontId < 0)
            {
                fonts.Add(font);
                fontId = fonts.Elements(S + "font").Count() - 1;
                fonts.SetAttributeValue("count", fontId + 1);
            }
            ChangeCellFormat(cell, xf => xf.SetAttributeValue("fontId", fontId), xf => xf.SetAttributeValue("applyFont", "1"));
        }

        private XElement? DefaultFont() => Font(0);

        private XElement? Font(int id) => _styles?.Root?.Element(S + "fonts")?.Elements(S + "font").ElementAtOrDefault(id);

        /// <summary>
        /// Give the cell a cell format like its current one with the change (an equal existing format is reused)
        /// </summary>
        private void ChangeCellFormat(XElement cell, Action<XElement> change, Action<XElement> apply)
        {
            XElement formats = CellFormats;
            var format = new XElement(CellFormat(cell) ?? new XElement(S + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0), new XAttribute("xfId", 0)));
            change(format);
            apply(format);
            List<XElement> all = formats.Elements(S + "xf").ToList();
            int index = all.FindIndex(f => XNode.DeepEquals(f, format));
            if (index < 0)
            {
                formats.Add(format);
                index = all.Count;
                formats.SetAttributeValue("count", index + 1);
            }
            if (index == 0)
                cell.Attribute("s")?.Remove();
            else
                cell.SetAttributeValue("s", index);
        }

        /// <summary>
        /// A number format shows dates if it has day, month, year or hour codes outside of quoted text and brackets (other than colors and
        /// conditions)
        /// </summary>
        internal static bool IsDateFormat(string format)
        {
            if (string.IsNullOrEmpty(format) || (format == "General") || (format == "@"))
                return false;
            bool quoted = false;
            bool bracket = false;
            foreach (char c in format)
            {
                if (c == '"') quoted = !quoted;
                else if (!quoted && (c == '[')) bracket = true;
                else if (!quoted && (c == ']')) bracket = false;
                else if (!quoted && !bracket && ("dmyhsDMYHS".IndexOf(c) >= 0)) return true;
            }
            return false;
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Addresses

        public static string ColumnName(int column)
        {
            string name = "";
            while (column > 0)
            {
                int remainder = (column - 1) % 26;
                name = (char)('A' + remainder) + name;
                column = (column - 1) / 26;
            }
            return name;
        }

        public static int ColumnIndex(string name)
        {
            int column = 0;
            foreach (char c in name.ToUpperInvariant())
            {
                if ((c < 'A') || (c > 'Z'))
                    throw new FormatException($"'{name}' is not a column");
                column = (column * 26) + (c - 'A' + 1);
            }
            return column;
        }

        /// <summary>
        /// "B2" / "$B$2" -> (2, 2)
        /// </summary>
        public static (int Row, int Column) ParseAddress(string address)
        {
            string text = address.Replace("$", "");
            int digits = 0;
            while ((digits < text.Length) && char.IsLetter(text[digits]))
                digits++;
            if ((digits == 0) || (digits == text.Length) || !int.TryParse(text.Substring(digits), NumberStyles.None, CultureInfo.InvariantCulture, out int row) || (row < 1))
                throw new FormatException($"'{address}' is not a cell address");
            return (row, ColumnIndex(text.Substring(0, digits)));
        }
    }
}
